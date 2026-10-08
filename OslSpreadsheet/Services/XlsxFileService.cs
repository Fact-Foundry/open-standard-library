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
                    Content = BuildWorksheet(sheet, styleIndexMap, isActiveSheet: sheet == workbook.Sheets[0])
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

            // Load styles.xml to detect date-formatted cells
            var dateXfIndices = new HashSet<int>();
            var stylesEntry = archive.GetEntry("xl/styles.xml");
            if (stylesEntry != null)
            {
                using var stylesStream = stylesEntry.Open();
                var stylesDoc = await Task.Run(() => XDocument.Load(stylesStream));
                dateXfIndices = GetDateXfIndices(stylesDoc, mainNs);
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

                // Shared formulas store the text once (on the master cell) and are reused by other cells with references shifted
                var sharedFormulas = new Dictionary<string, (string formula, int row, int col)>();

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
                            case "e": // error value such as #N/A
                                value = rawValue;
                                break;
                            case "b": // boolean
                                value = rawValue == "1" ? "true" : "false";
                                valueType = CellValueType.Boolean;
                                break;
                            default: // number
                                value = rawValue;
                                if (!string.IsNullOrEmpty(value))
                                {
                                    var styleIdx = int.TryParse(cellEl.Attribute("s")?.Value, out int s) ? s : -1;
                                    if (styleIdx >= 0 && dateXfIndices.Contains(styleIdx) && double.TryParse(rawValue, out double serial))
                                    {
                                        valueType = CellValueType.DateTime;
                                        value = System.DateTime.FromOADate(serial).ToString("yyyy-MM-ddTHH:mm:ss");
                                    }
                                    else
                                    {
                                        valueType = CellValueType.Float;
                                    }
                                }
                                break;
                        }

                        var oCell = sheet.AddCell(rowNum, colNum, value);
                        oCell.ValueType = valueType;
                        oCell.Formula = ReadFormula(cellEl.Element(mainNs + "f"), rowNum, colNum, sharedFormulas);
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

                if ((sheet.AutoFilterRange?.StartRow == 1) || sheet.FreezeRows >= 1)
                    sheet.HasHeaderRow = true;
            }

            return workbook;
        }

        private static string? ReadFormula(XElement? f, int row, int col, Dictionary<string, (string formula, int row, int col)> sharedFormulas)
        {
            if (f == null)
                return null;

            var text = f.Value;
            var sharedIndex = f.Attribute("t")?.Value == "shared" ? f.Attribute("si")?.Value : null;

            if (sharedIndex != null)
            {
                if (!string.IsNullOrEmpty(text))
                    sharedFormulas[sharedIndex] = (text, row, col);
                else if (sharedFormulas.TryGetValue(sharedIndex, out var master))
                    text = FormulaTranslator.Shift(master.formula, row - master.row, col - master.col);
            }

            return string.IsNullOrEmpty(text) ? null : FormulaTranslator.FromXlsx(text);
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
            // The library doesn't calculate formulas, so ask the application to compute them when the file opens
            if (workbook.Sheets.Any(s => s.Cells.Any(c => FormulaTranslator.Normalize(c.Formula) != null)))
                sb.Append("<calcPr fullCalcOnLoad=\"1\"/>");
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

        private const int DateTimeNumFmtId = 164;
        private const string DateTimeFormatCode = "yyyy-mm-dd hh:mm:ss";
        private const int DateOnlyNumFmtId = 165;
        private const string DateOnlyFormatCode = "yyyy-mm-dd";

        private static bool IsDateOnly(string value) =>
            System.DateTime.TryParse(value, out var dt) && dt.TimeOfDay == TimeSpan.Zero && !value.Contains('T');

        private static readonly HashSet<int> BuiltInDateNumFmtIds =
            new(Enumerable.Range(14, 9)
                .Concat(Enumerable.Range(27, 10))
                .Concat(Enumerable.Range(45, 3))
                .Concat(Enumerable.Range(50, 9)));

        private static (byte[] stylesXml, Dictionary<string, int> styleIndexMap) BuildStylesForWorkbook(oWorkbook workbook)
        {
            bool hasDateCells = workbook.Sheets.Any(s => s.Cells.Any(c => c.ValueType == CellValueType.DateTime));

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
            var xfKeys = new Dictionary<string, int> { ["0|0|0|False|0"] = 0 };

            var styleIndexMap = new Dictionary<string, int>();

            var uniqueEntries = new Dictionary<string, (CellStyle style, int numFmtId)>();
            foreach (var sheet in workbook.Sheets)
                foreach (var cell in sheet.Cells)
                {
                    int numFmtId = cell.ValueType == CellValueType.DateTime
                        ? (IsDateOnly(cell.Value) ? DateOnlyNumFmtId : DateTimeNumFmtId)
                        : 0;
                    var style = cell.Style ?? new CellStyle();
                    var mapKey = numFmtId > 0 ? $"dt{numFmtId}|{GetStyleKey(style)}" : GetStyleKey(style);

                    if (cell.Style != null || numFmtId > 0)
                        uniqueEntries.TryAdd(mapKey, (style, numFmtId));
                }

            foreach (var (key, (style, numFmtId)) in uniqueEntries)
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

                var xfk = $"{fontId}|{fillId}|{borderId}|{style.WrapText}|{numFmtId}";
                if (!xfKeys.TryGetValue(xfk, out int xfId))
                {
                    xfId = xfs.Count;
                    var xfSb = new StringBuilder($"<xf numFmtId=\"{numFmtId}\" fontId=\"{fontId}\" fillId=\"{fillId}\" borderId=\"{borderId}\" xfId=\"0\"");
                    if (numFmtId > 0) xfSb.Append(" applyNumberFormat=\"1\"");
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
            if (hasDateCells)
            {
                var usedFmtIds = uniqueEntries.Values.Select(e => e.numFmtId).Where(id => id > 0).Distinct().ToList();
                sb.Append($"<numFmts count=\"{usedFmtIds.Count}\">");
                if (usedFmtIds.Contains(DateTimeNumFmtId))
                    sb.Append($"<numFmt numFmtId=\"{DateTimeNumFmtId}\" formatCode=\"{DateTimeFormatCode}\"/>");
                if (usedFmtIds.Contains(DateOnlyNumFmtId))
                    sb.Append($"<numFmt numFmtId=\"{DateOnlyNumFmtId}\" formatCode=\"{DateOnlyFormatCode}\"/>");
                sb.Append("</numFmts>");
            }
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

        private static HashSet<int> GetDateXfIndices(XDocument stylesDoc, XNamespace ns)
        {
            var dateNumFmtIds = new HashSet<int>(BuiltInDateNumFmtIds);

            foreach (var nf in stylesDoc.Descendants(ns + "numFmt"))
            {
                var id = int.TryParse(nf.Attribute("numFmtId")?.Value, out int nfId) ? nfId : 0;
                var code = nf.Attribute("formatCode")?.Value ?? "";
                if (code.Contains('y') || code.Contains('d') || code.Contains('h'))
                    dateNumFmtIds.Add(id);
            }

            var result = new HashSet<int>();
            var xfs = stylesDoc.Descendants(ns + "cellXfs").Elements(ns + "xf").ToList();
            for (int i = 0; i < xfs.Count; i++)
            {
                var numFmtId = int.TryParse(xfs[i].Attribute("numFmtId")?.Value, out int nfId) ? nfId : 0;
                if (dateNumFmtIds.Contains(numFmtId))
                    result.Add(i);
            }

            return result;
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

        private static byte[] BuildWorksheet(oSpreadsheet sheet, Dictionary<string, int> styleIndexMap, bool isActiveSheet)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

            if (sheet.FreezeRows > 0 || sheet.FreezeColumns > 0)
            {
                var topLeftCell = $"{ColumnLetter(sheet.FreezeColumns + 1)}{sheet.FreezeRows + 1}";
                // Only one sheet may be selected; Excel treats multiple selected sheets as a group and applies edits to all of them
                sb.Append(isActiveSheet ? "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">" : "<sheetViews><sheetView workbookViewId=\"0\">");
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

            // Group cells by row in one pass rather than scanning the cell list once per row
            var rows = sheet.Cells.GroupBy(c => c.Row).OrderBy(g => g.Key);

            foreach (var rowGroup in rows)
            {
                var r = rowGroup.Key;
                sb.Append($"<row r=\"{r}\">");
                foreach (var cell in rowGroup.OrderBy(c => c.Column))
                {
                    var cellRef = $"{ColumnLetter(cell.Column)}{cell.Row}";
                    var styleAttr = "";
                    var visualKey = cell.Style != null ? GetStyleKey(cell.Style) : null;

                    if (cell.ValueType == CellValueType.DateTime)
                    {
                        int fmtId = IsDateOnly(cell.Value) ? DateOnlyNumFmtId : DateTimeNumFmtId;
                        var dateKey = visualKey != null ? $"dt{fmtId}|{visualKey}" : $"dt{fmtId}|{GetStyleKey(new CellStyle())}";
                        if (styleIndexMap.TryGetValue(dateKey, out int si))
                            styleAttr = $" s=\"{si}\"";
                    }
                    else if (visualKey != null)
                    {
                        if (styleIndexMap.TryGetValue(visualKey, out int si) && si > 0)
                            styleAttr = $" s=\"{si}\"";
                    }

                    if (FormulaTranslator.Normalize(cell.Formula) is string formula)
                        sb.Append(BuildFormulaCell(cell, cellRef, styleAttr, formula));
                    else if (cell.ValueType == CellValueType.DateTime && System.DateTime.TryParse(cell.Value, out var dt))
                        sb.Append($"<c r=\"{cellRef}\"{styleAttr}><v>{dt.ToOADate()}</v></c>");
                    else if (cell.ValueType == CellValueType.Float || cell.ValueType == CellValueType.Int64)
                        sb.Append($"<c r=\"{cellRef}\"{styleAttr}><v>{SecurityElement.Escape(cell.Value)}</v></c>");
                    else if (cell.ValueType == CellValueType.Boolean)
                        sb.Append($"<c r=\"{cellRef}\"{styleAttr} t=\"b\"><v>{(cell.Value.Equals("true", StringComparison.OrdinalIgnoreCase) ? "1" : "0")}</v></c>");
                    else
                        sb.Append($"<c r=\"{cellRef}\"{styleAttr} t=\"inlineStr\"><is><t xml:space=\"preserve\">{SecurityElement.Escape(cell.Value)}</t></is></c>");
                }
                sb.Append("</row>");
            }

            sb.Append("</sheetData>");

            if (sheet.AutoFilterRange is var (sr, sc, er, ec))
                sb.Append($"<autoFilter ref=\"{ColumnLetter(sc)}{sr}:{ColumnLetter(ec)}{er}\"/>");

            sb.Append("</worksheet>");
            return Utf8(sb.ToString());
        }

        /// <summary>
        /// Builds a formula cell. The cell's Value, if set, is written as the cached result; otherwise no result is stored.
        /// </summary>
        private static string BuildFormulaCell(oCell cell, string cellRef, string styleAttr, string formula)
        {
            var f = $"<f>{SecurityElement.Escape(FormulaTranslator.ToXlsx(formula))}</f>";

            if (string.IsNullOrEmpty(cell.Value))
                return $"<c r=\"{cellRef}\"{styleAttr}>{f}</c>";

            return cell.ValueType switch
            {
                CellValueType.Float or CellValueType.Int64 when double.TryParse(cell.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out _)
                    => $"<c r=\"{cellRef}\"{styleAttr}>{f}<v>{SecurityElement.Escape(cell.Value)}</v></c>",
                CellValueType.DateTime when System.DateTime.TryParse(cell.Value, out var dt)
                    => $"<c r=\"{cellRef}\"{styleAttr}>{f}<v>{dt.ToOADate()}</v></c>",
                CellValueType.Boolean
                    => $"<c r=\"{cellRef}\"{styleAttr} t=\"b\">{f}<v>{(cell.Value.Equals("true", StringComparison.OrdinalIgnoreCase) ? "1" : "0")}</v></c>",
                _ => $"<c r=\"{cellRef}\"{styleAttr} t=\"str\">{f}<v>{SecurityElement.Escape(cell.Value)}</v></c>"
            };
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
