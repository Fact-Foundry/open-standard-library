using OslSpreadsheet.Models;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

namespace OslSpreadsheet.Services
{
    internal class XlsxFileService : IFileService
    {
        private bool disposedValue;

        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }

        public async Task<byte[]> GenerateFileAsync(oWorkbook workbook)
        {
            var (stylesXml, styleIndexMap) = BuildStylesForWorkbook(workbook);

            var files = new List<InMemoryFile>
            {
                new InMemoryFile { FileName = "[Content_Types].xml", Content = BuildContentTypes(workbook) },
                new InMemoryFile { FileName = "_rels/.rels",          Content = BuildRootRels() },
                new InMemoryFile { FileName = "xl/workbook.xml",      Content = BuildWorkbook(workbook) },
                new InMemoryFile { FileName = "xl/_rels/workbook.xml.rels", Content = BuildWorkbookRels(workbook) },
                new InMemoryFile { FileName = "xl/styles.xml",        Content = stylesXml },
            };

            foreach (var sheet in workbook.Sheets)
            {
                files.Add(new InMemoryFile
                {
                    FileName = $"xl/worksheets/sheet{sheet.Index}.xml",
                    Content = BuildWorksheet(sheet, styleIndexMap)
                });
            }

            await using MemoryStream archiveStream = new();
            using (ZipArchive archive = new(archiveStream, ZipArchiveMode.Create, true))
            {
                foreach (var file in files)
                {
                    var entry = archive.CreateEntry(file.FileName, CompressionLevel.Fastest);
                    using var stream = entry.Open();
                    stream.Write(file.Content, 0, file.Content.Length);
                }
            }

            return archiveStream.ToArray();
        }

        public async Task<oWorkbook> GenerateModel(byte[] file)
        {
            oWorkbook workbook = new();

            using var ms = new MemoryStream(file);
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

            XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace rNs    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace relsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            // Load shared strings if present
            var sharedStrings = new List<string>();
            var ssEntry = archive.GetEntry("xl/sharedStrings.xml");
            if (ssEntry != null)
            {
                using var ssStream = ssEntry.Open();
                var ssDoc = await Task.Run(() => XDocument.Load(ssStream));
                sharedStrings = ssDoc.Descendants(mainNs + "si")
                    .Select(si => string.Concat(si.Descendants(mainNs + "t").Select(t => t.Value)))
                    .ToList();
            }

            // Load workbook.xml
            var wbEntry = archive.GetEntry("xl/workbook.xml")
                ?? throw new InvalidOperationException("Invalid XLSX file: missing xl/workbook.xml");

            XDocument wbDoc;
            using (var wbStream = wbEntry.Open())
                wbDoc = await Task.Run(() => XDocument.Load(wbStream));

            // Load workbook.xml.rels — maps rId to sheet file path
            var relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
                ?? throw new InvalidOperationException("Invalid XLSX file: missing xl/_rels/workbook.xml.rels");

            XDocument relsDoc;
            using (var relsStream = relsEntry.Open())
                relsDoc = await Task.Run(() => XDocument.Load(relsStream));

            var relationships = relsDoc.Descendants(relsNs + "Relationship")
                .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);

            foreach (var sheetEl in wbDoc.Descendants(mainNs + "sheet"))
            {
                var sheetName = sheetEl.Attribute("name")?.Value ?? "Sheet";
                var rId = sheetEl.Attribute(rNs + "id")?.Value;
                if (rId == null || !relationships.TryGetValue(rId, out var target)) continue;

                // Target paths are relative to xl/
                var sheetPath = target.StartsWith('/') ? target.TrimStart('/') : $"xl/{target}";
                var sheetEntry = archive.GetEntry(sheetPath);
                if (sheetEntry == null) continue;

                var sheet = workbook.AddSheet(sheetName);

                XDocument sheetDoc;
                using (var sheetStream = sheetEntry.Open())
                    sheetDoc = await Task.Run(() => XDocument.Load(sheetStream));

                var pane = sheetDoc.Descendants(mainNs + "pane").FirstOrDefault();
                if (pane?.Attribute("state")?.Value == "frozen")
                {
                    if (int.TryParse(pane.Attribute("ySplit")?.Value, out int freezeRows))
                        sheet.FreezeRows = freezeRows;
                    if (int.TryParse(pane.Attribute("xSplit")?.Value, out int freezeCols))
                        sheet.FreezeColumns = freezeCols;
                }

                foreach (var rowEl in sheetDoc.Descendants(mainNs + "row"))
                {
                    foreach (var cellEl in rowEl.Elements(mainNs + "c"))
                    {
                        var cellRef = cellEl.Attribute("r")?.Value;
                        if (cellRef == null) continue;

                        var (rowNum, colNum) = ParseCellRef(cellRef);
                        var cellType = cellEl.Attribute("t")?.Value;
                        var rawValue = cellEl.Element(mainNs + "v")?.Value ?? "";

                        string value;
                        CellValueType valueType = CellValueType.String;

                        switch (cellType)
                        {
                            case "s": // shared string index
                                var idx = int.TryParse(rawValue, out int si) ? si : 0;
                                value = idx < sharedStrings.Count ? sharedStrings[idx] : "";
                                break;
                            case "inlineStr":
                                value = string.Concat(cellEl.Descendants(mainNs + "t").Select(t => t.Value));
                                break;
                            case "str": // formula string result
                                value = rawValue;
                                break;
                            case "b": // boolean
                                value = rawValue == "1" ? "true" : "false";
                                valueType = CellValueType.Boolean;
                                break;
                            default: // number
                                value = rawValue;
                                if (!string.IsNullOrEmpty(value))
                                    valueType = CellValueType.Float;
                                break;
                        }

                        var oCell = sheet.AddCell(rowNum, colNum, value);
                        oCell.ValueType = valueType;
                    }
                }

                var autoFilter = sheetDoc.Descendants(mainNs + "autoFilter").FirstOrDefault();
                if (autoFilter?.Attribute("ref")?.Value is string afRef)
                {
                    var parts = afRef.Split(':');
                    if (parts.Length == 2)
                    {
                        var (startRow, startCol) = ParseCellRef(parts[0]);
                        var (endRow, endCol) = ParseCellRef(parts[1]);
                        sheet.AutoFilterRange = (startRow, startCol, endRow, endCol);
                    }
                }
            }

            return workbook;
        }

        private static (int row, int col) ParseCellRef(string cellRef)
        {
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;

            int col = 0;
            foreach (char c in cellRef[..i])
                col = col * 26 + (c - 'A' + 1);

            return (int.Parse(cellRef[i..]), col);
        }

        private static byte[] Utf8(string xml) => Encoding.UTF8.GetBytes(xml);

        private static byte[] BuildContentTypes(oWorkbook workbook)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            foreach (var sheet in workbook.Sheets)
                sb.Append($"<Override PartName=\"/xl/worksheets/sheet{sheet.Index}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            sb.Append("</Types>");
            return Utf8(sb.ToString());
        }

        private static byte[] BuildRootRels() => Utf8(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>");

        private static byte[] BuildWorkbook(oWorkbook workbook)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<sheets>");
            foreach (var sheet in workbook.Sheets)
                sb.Append($"<sheet name=\"{SecurityElement.Escape(sheet.SheetName)}\" sheetId=\"{sheet.Index}\" r:id=\"rId{sheet.Index}\"/>");
            sb.Append("</sheets>");
            sb.Append("</workbook>");
            return Utf8(sb.ToString());
        }

        private static byte[] BuildWorkbookRels(oWorkbook workbook)
        {
            var stylesId = workbook.Sheets.Any() ? workbook.Sheets.Max(x => x.Index) + 1 : 1;

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var sheet in workbook.Sheets)
                sb.Append($"<Relationship Id=\"rId{sheet.Index}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{sheet.Index}.xml\"/>");
            sb.Append($"<Relationship Id=\"rId{stylesId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            sb.Append("</Relationships>");
            return Utf8(sb.ToString());
        }

        private static (byte[] stylesXml, Dictionary<string, int> styleIndexMap) BuildStylesForWorkbook(oWorkbook workbook)
        {
            var defaultFontKey = GetFontKey(new CellStyle());
            var fonts = new List<string> { "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>" };
            var fontKeys = new Dictionary<string, int> { [defaultFontKey] = 0 };

            var fills = new List<string>
            {
                "<fill><patternFill patternType=\"none\"/></fill>",
                "<fill><patternFill patternType=\"gray125\"/></fill>"
            };
            var fillKeys = new Dictionary<string, int> { ["none"] = 0 };

            var defaultBorderKey = GetBorderKey(new CellStyle());
            var borders = new List<string> { "<border><left/><right/><top/><bottom/><diagonal/></border>" };
            var borderKeys = new Dictionary<string, int> { [defaultBorderKey] = 0 };

            var xfs = new List<string> { "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" };
            var xfKeys = new Dictionary<string, int> { ["0|0|0"] = 0 };

            var styleIndexMap = new Dictionary<string, int>();

            var uniqueStyles = new Dictionary<string, CellStyle>();
            foreach (var sheet in workbook.Sheets)
                foreach (var cell in sheet.Cells)
                    if (cell.Style != null)
                    {
                        var key = GetStyleKey(cell.Style);
                        uniqueStyles.TryAdd(key, cell.Style);
                    }

            foreach (var (key, style) in uniqueStyles)
            {
                var fk = GetFontKey(style);
                if (!fontKeys.TryGetValue(fk, out int fontId))
                {
                    fontId = fonts.Count;
                    fonts.Add(BuildFontXml(style));
                    fontKeys[fk] = fontId;
                }

                var flk = style.BackgroundColor ?? "none";
                if (!fillKeys.TryGetValue(flk, out int fillId))
                {
                    fillId = fills.Count;
                    fills.Add($"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{ToArgb(style.BackgroundColor!)}\"/></patternFill></fill>");
                    fillKeys[flk] = fillId;
                }

                var bk = GetBorderKey(style);
                if (!borderKeys.TryGetValue(bk, out int borderId))
                {
                    borderId = borders.Count;
                    borders.Add(BuildBorderXml(style));
                    borderKeys[bk] = borderId;
                }

                var xfk = $"{fontId}|{fillId}|{borderId}|{style.WrapText}";
                if (!xfKeys.TryGetValue(xfk, out int xfId))
                {
                    xfId = xfs.Count;
                    var xfSb = new StringBuilder($"<xf numFmtId=\"0\" fontId=\"{fontId}\" fillId=\"{fillId}\" borderId=\"{borderId}\" xfId=\"0\"");
                    if (fontId > 0) xfSb.Append(" applyFont=\"1\"");
                    if (fillId > 0) xfSb.Append(" applyFill=\"1\"");
                    if (borderId > 0) xfSb.Append(" applyBorder=\"1\"");
                    if (style.WrapText) xfSb.Append(" applyAlignment=\"1\"");
                    if (style.WrapText)
                    {
                        xfSb.Append("><alignment wrapText=\"1\"/></xf>");
                    }
                    else
                    {
                        xfSb.Append("/>");
                    }
                    xfs.Add(xfSb.ToString());
                    xfKeys[xfk] = xfId;
                }

                styleIndexMap[key] = xfId;
            }

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append($"<fonts count=\"{fonts.Count}\">");
            foreach (var f in fonts) sb.Append(f);
            sb.Append("</fonts>");
            sb.Append($"<fills count=\"{fills.Count}\">");
            foreach (var f in fills) sb.Append(f);
            sb.Append("</fills>");
            sb.Append($"<borders count=\"{borders.Count}\">");
            foreach (var b in borders) sb.Append(b);
            sb.Append("</borders>");
            sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
            sb.Append($"<cellXfs count=\"{xfs.Count}\">");
            foreach (var x in xfs) sb.Append(x);
            sb.Append("</cellXfs>");
            sb.Append("</styleSheet>");

            return (Utf8(sb.ToString()), styleIndexMap);
        }

        private static string GetStyleKey(CellStyle s) =>
            $"{s.Bold}|{s.Italic}|{s.Underline}|{s.FontColor}|{s.BackgroundColor}|{s.FontName}|{s.FontSize}|{s.WrapText}|{EdgeKey(s.BorderTop)}|{EdgeKey(s.BorderBottom)}|{EdgeKey(s.BorderLeft)}|{EdgeKey(s.BorderRight)}";

        private static string GetFontKey(CellStyle s) =>
            $"{s.Bold}|{s.Italic}|{s.Underline}|{s.FontColor}|{s.FontName}|{s.FontSize}";

        private static string GetBorderKey(CellStyle s) =>
            $"{EdgeKey(s.BorderTop)}|{EdgeKey(s.BorderBottom)}|{EdgeKey(s.BorderLeft)}|{EdgeKey(s.BorderRight)}";

        private static string EdgeKey(CellBorder? b) =>
            b == null ? "" : $"{b.Style}:{b.Color}";

        private static string ToArgb(string hex) => "FF" + hex.TrimStart('#');

        private static string BuildFontXml(CellStyle s)
        {
            var sb = new StringBuilder("<font>");
            if (s.Bold) sb.Append("<b/>");
            if (s.Italic) sb.Append("<i/>");
            if (s.Underline) sb.Append("<u/>");
            sb.Append($"<sz val=\"{s.FontSize ?? 11}\"/>");
            if (s.FontColor != null)
                sb.Append($"<color rgb=\"{ToArgb(s.FontColor)}\"/>");
            sb.Append($"<name val=\"{SecurityElement.Escape(s.FontName ?? "Calibri")}\"/>");
            sb.Append("</font>");
            return sb.ToString();
        }

        private static string BuildBorderXml(CellStyle s)
        {
            var sb = new StringBuilder("<border>");
            sb.Append(BuildBorderEdgeXml("left", s.BorderLeft));
            sb.Append(BuildBorderEdgeXml("right", s.BorderRight));
            sb.Append(BuildBorderEdgeXml("top", s.BorderTop));
            sb.Append(BuildBorderEdgeXml("bottom", s.BorderBottom));
            sb.Append("<diagonal/>");
            sb.Append("</border>");
            return sb.ToString();
        }

        private static string BuildBorderEdgeXml(string edge, CellBorder? b)
        {
            if (b == null || b.Style == BorderStyle.None)
                return $"<{edge}/>";
            var style = b.Style.ToString().ToLowerInvariant();
            if (b.Color != null)
                return $"<{edge} style=\"{style}\"><color rgb=\"{ToArgb(b.Color)}\"/></{edge}>";
            return $"<{edge} style=\"{style}\"/>";
        }

        private static byte[] BuildWorksheet(oSpreadsheet sheet, Dictionary<string, int> styleIndexMap)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

            if (sheet.FreezeRows > 0 || sheet.FreezeColumns > 0)
            {
                var topLeftCell = $"{ColumnLetter(sheet.FreezeColumns + 1)}{sheet.FreezeRows + 1}";
                sb.Append("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">");
                sb.Append($"<pane");
                if (sheet.FreezeColumns > 0) sb.Append($" xSplit=\"{sheet.FreezeColumns}\"");
                if (sheet.FreezeRows > 0) sb.Append($" ySplit=\"{sheet.FreezeRows}\"");
                sb.Append($" topLeftCell=\"{topLeftCell}\" activePane=\"bottomRight\" state=\"frozen\"/>");
                sb.Append("</sheetView></sheetViews>");
            }

            if (sheet.ColumnWidths.Any())
            {
                sb.Append("<cols>");
                foreach (var (col, width) in sheet.ColumnWidths.OrderBy(x => x.Key))
                    sb.Append($"<col min=\"{col}\" max=\"{col}\" width=\"{width}\" customWidth=\"1\"/>");
                sb.Append("</cols>");
            }

            sb.Append("<sheetData>");

            for (int r = 1; r <= sheet.RowCount; r++)
            {
                var rowCells = sheet.GetRow(r);
                if (!rowCells.Any()) continue;

                sb.Append($"<row r=\"{r}\">");
                foreach (var cell in rowCells)
                {
                    var cellRef = $"{ColumnLetter(cell.Column)}{cell.Row}";
                    var styleAttr = "";
                    if (cell.Style != null)
                    {
                        var key = GetStyleKey(cell.Style);
                        if (styleIndexMap.TryGetValue(key, out int si) && si > 0)
                            styleAttr = $" s=\"{si}\"";
                    }

                    if (cell.ValueType == CellValueType.Float)
                        sb.Append($"<c r=\"{cellRef}\"{styleAttr}><v>{SecurityElement.Escape(cell.Value)}</v></c>");
                    else if (cell.ValueType == CellValueType.Boolean)
                        sb.Append($"<c r=\"{cellRef}\"{styleAttr} t=\"b\"><v>{(cell.Value.Equals("true", StringComparison.OrdinalIgnoreCase) ? "1" : "0")}</v></c>");
                    else
                        sb.Append($"<c r=\"{cellRef}\"{styleAttr} t=\"inlineStr\"><is><t>{SecurityElement.Escape(cell.Value)}</t></is></c>");
                }
                sb.Append("</row>");
            }

            sb.Append("</sheetData>");

            if (sheet.AutoFilterRange is var (sr, sc, er, ec))
                sb.Append($"<autoFilter ref=\"{ColumnLetter(sc)}{sr}:{ColumnLetter(ec)}{er}\"/>");

            sb.Append("</worksheet>");
            return Utf8(sb.ToString());
        }

        private static string ColumnLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                col--;
                result = (char)('A' + col % 26) + result;
                col /= 26;
            }
            return result;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
                disposedValue = true;
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
