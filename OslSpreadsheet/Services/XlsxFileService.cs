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
            var files = new List<InMemoryFile>
            {
                new InMemoryFile { FileName = "[Content_Types].xml", Content = BuildContentTypes(workbook) },
                new InMemoryFile { FileName = "_rels/.rels",          Content = BuildRootRels() },
                new InMemoryFile { FileName = "xl/workbook.xml",      Content = BuildWorkbook(workbook) },
                new InMemoryFile { FileName = "xl/_rels/workbook.xml.rels", Content = BuildWorkbookRels(workbook) },
                new InMemoryFile { FileName = "xl/styles.xml",        Content = BuildStyles() },
            };

            foreach (var sheet in workbook.Sheets)
            {
                files.Add(new InMemoryFile
                {
                    FileName = $"xl/worksheets/sheet{sheet.Index}.xml",
                    Content = BuildWorksheet(sheet)
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

        private static byte[] BuildStyles() => Utf8(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
            "<fills count=\"2\">" +
                "<fill><patternFill patternType=\"none\"/></fill>" +
                "<fill><patternFill patternType=\"gray125\"/></fill>" +
            "</fills>" +
            "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>" +
            "</styleSheet>");

        private static byte[] BuildWorksheet(oSpreadsheet sheet)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<sheetData>");

            for (int r = 1; r <= sheet.RowCount; r++)
            {
                var rowCells = sheet.GetRow(r);
                if (!rowCells.Any()) continue;

                sb.Append($"<row r=\"{r}\">");
                foreach (var cell in rowCells)
                {
                    var cellRef = $"{ColumnLetter(cell.Column)}{cell.Row}";
                    if (cell.ValueType == CellValueType.Float)
                        sb.Append($"<c r=\"{cellRef}\"><v>{SecurityElement.Escape(cell.Value)}</v></c>");
                    else
                        sb.Append($"<c r=\"{cellRef}\" t=\"inlineStr\"><is><t>{SecurityElement.Escape(cell.Value)}</t></is></c>");
                }
                sb.Append("</row>");
            }

            sb.Append("</sheetData>");
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
