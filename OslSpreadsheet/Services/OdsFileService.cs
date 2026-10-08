using OslSpreadsheet.Models;
using OslSpreadsheet.Models.Files.ods;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;

namespace OslSpreadsheet.Services
{
    internal class OdsFileService : IFileService
    {
        private bool disposedValue;

        public async ValueTask DisposeAsync()
        {
            await Task.Run(() => Dispose());
        }

        public async Task<byte[]> GenerateFileAsync(oWorkbook workbook)
        {
            byte[] output;

            try
            {
                var (content, hasDateOnly, hasDateTime) = GenerateContentFile(workbook);
                var meta = GenerateMetaFile(workbook);
                var style = GenerateStyleFile(workbook);

                var manifest = new ODManifest();
                bool hasSettings = workbook.Sheets.Any(s => s.FreezeRows > 0 || s.FreezeColumns > 0);
                if (hasSettings)
                    manifest.fileEntries.Add(new ODManifest.FileEntry() { FullPath = "settings.xml", MediaType = "text/xml" });

                // Add models to a file list for compression
                // mimetype must be first and uncompressed per ODS spec
                List<InMemoryFile> files = new()
                {
                    new InMemoryFile()
                    {
                        FileName = "mimetype",
                        Content = Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.spreadsheet"),
                        Store = true
                    },
                    new InMemoryFile()
                    {
                        FileName = "META-INF/manifest.xml",
                        Content = await XmlService.ConvertToXmlAsync(manifest)
                    },
                    new InMemoryFile()
                    {
                        FileName = "content.xml",
                        Content = InjectOdsDateStyles(await XmlService.ConvertToXmlAsync(content), hasDateOnly, hasDateTime)
                    },
                    new InMemoryFile()
                    {
                        FileName = "meta.xml",
                        Content = await XmlService.ConvertToXmlAsync(meta)
                    },
                    new InMemoryFile()
                    {
                        FileName = "styles.xml",
                        Content = await XmlService.ConvertToXmlAsync(style)
                    }
                };

                if (hasSettings)
                {
                    files.Add(new InMemoryFile()
                    {
                        FileName = "settings.xml",
                        Content = BuildSettingsFile(workbook)
                    });
                }

                output = await ZipService.GenerateZipAsync(files);
            }
            catch
            {
                throw new Exception("There was an issue compressing the file.");
            }

            return output;
        }

        public async Task<oWorkbook> GenerateModel(byte[] file)
        {
            oWorkbook workbook = new();

            using var ms = new MemoryStream(file);
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

            var contentEntry = archive.GetEntry("content.xml")
                ?? throw new InvalidOperationException("Invalid ODS file: missing content.xml");

            XDocument doc;
            using (var contentStream = contentEntry.Open())
                doc = await Task.Run(() => XDocument.Load(contentStream));

            XNamespace tableNs  = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
            XNamespace officeNs = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XNamespace textNs   = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

            foreach (var table in doc.Descendants(tableNs + "table"))
            {
                var sheetName = table.Attribute(tableNs + "name")?.Value ?? "Sheet";
                var sheet = workbook.AddSheet(sheetName);

                int rowIndex = 0;

                foreach (var tableRow in GetTableRows(table, tableNs))
                {
                    int rowsRepeated = int.TryParse(tableRow.Attribute(tableNs + "number-rows-repeated")?.Value, out int rr) ? rr : 1;

                    var rowData = new List<(int col, string value, CellValueType type, string? formula)>();
                    int colIndex = 0;

                    foreach (var cell in tableRow.Elements(tableNs + "table-cell"))
                    {
                        int colsRepeated = int.TryParse(cell.Attribute(tableNs + "number-columns-repeated")?.Value, out int cr) ? cr : 1;

                        var valueType  = cell.Attribute(officeNs + "value-type")?.Value;
                        var textValue  = cell.Element(textNs + "p")?.Value;
                        var numericValue = cell.Attribute(officeNs + "value")?.Value;
                        var booleanValue = cell.Attribute(officeNs + "boolean-value")?.Value;
                        var formulaValue = cell.Attribute(tableNs + "formula")?.Value;
                        var formula = string.IsNullOrEmpty(formulaValue) ? null : FormulaTranslator.FromOpenFormula(formulaValue);
                        bool hasContent = valueType != null || textValue != null || formula != null;

                        if (hasContent)
                        {
                            CellValueType cellType;
                            string cellValue;

                            if (valueType == "boolean")
                            {
                                cellType = CellValueType.Boolean;
                                cellValue = booleanValue ?? textValue ?? "false";
                            }
                            else if (valueType == "date")
                            {
                                cellType = CellValueType.DateTime;
                                var dateValue = cell.Attribute(officeNs + "date-value")?.Value;
                                cellValue = dateValue ?? textValue ?? "";
                            }
                            else if (valueType == "float")
                            {
                                cellType = CellValueType.Float;
                                cellValue = textValue ?? numericValue ?? "";
                            }
                            else
                            {
                                cellType = CellValueType.String;
                                cellValue = textValue ?? numericValue ?? "";
                            }

                            for (int i = 0; i < colsRepeated; i++)
                            {
                                colIndex++;
                                rowData.Add((colIndex, cellValue, cellType, formula));
                            }
                        }
                        else
                        {
                            colIndex += colsRepeated;
                        }
                    }

                    if (rowData.Count == 0)
                    {
                        rowIndex += rowsRepeated;
                        continue;
                    }

                    for (int r = 0; r < rowsRepeated; r++)
                    {
                        rowIndex++;
                        foreach (var (col, value, type, formula) in rowData)
                        {
                            var oCell = sheet.AddCell(rowIndex, col, value);
                            oCell.ValueType = type;
                            oCell.Formula = formula;
                        }
                    }
                }
            }

            var settingsEntry = archive.GetEntry("settings.xml");
            if (settingsEntry != null)
            {
                XDocument settingsDoc;
                using (var settingsStream = settingsEntry.Open())
                    settingsDoc = await Task.Run(() => XDocument.Load(settingsStream));

                XNamespace configNs = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
                foreach (var entry in settingsDoc.Descendants(configNs + "config-item-map-named")
                    .Where(n => n.Attribute(configNs + "name")?.Value == "Tables")
                    .SelectMany(n => n.Elements(configNs + "config-item-map-entry")))
                {
                    var tableName = entry.Attribute(configNs + "name")?.Value;
                    var sheet = workbook.Sheets.FirstOrDefault(s => s.SheetName == tableName);
                    if (sheet == null) continue;

                    var items = entry.Elements(configNs + "config-item")
                        .ToDictionary(e => e.Attribute(configNs + "name")?.Value ?? "", e => e.Value);

                    if (items.TryGetValue("VerticalSplitPosition", out var vsp) && int.TryParse(vsp, out int freezeRows))
                        sheet.FreezeRows = freezeRows;
                    if (items.TryGetValue("HorizontalSplitPosition", out var hsp) && int.TryParse(hsp, out int freezeCols))
                        sheet.FreezeColumns = freezeCols;
                }
            }

            foreach (var dbRange in doc.Descendants(tableNs + "database-range"))
            {
                var displayButtons = dbRange.Attribute(tableNs + "display-filter-buttons")?.Value;
                var targetAddr = dbRange.Attribute(tableNs + "target-range-address")?.Value;
                if (displayButtons != "true" || targetAddr == null) continue;

                var parts = targetAddr.Split(':');
                if (parts.Length != 2) continue;

                var (sheetName1, startRow, startCol) = ParseOdsCellAddress(parts[0]);
                var (_, endRow, endCol) = ParseOdsCellAddress(parts[1]);

                var sheet = workbook.Sheets.FirstOrDefault(s => s.SheetName == sheetName1);
                if (sheet != null)
                    sheet.AutoFilterRange = (startRow, startCol, endRow, endCol);
            }

            foreach (var sheet in workbook.Sheets)
            {
                if ((sheet.AutoFilterRange?.StartRow == 1) || sheet.FreezeRows >= 1)
                    sheet.HasHeaderRow = true;
            }

            return workbook;
        }

        private (ODContent content, bool hasDateOnly, bool hasDateTime) GenerateContentFile(oWorkbook workbook)
        {
            var file = new ODContent();

            var cellStyleMap = BuildOdsCellStyles(workbook, file, out bool hasDateOnly, out bool hasDateTime);

            foreach (var s in workbook.Sheets)
            {
                var masterPageName = string.Format("mp{0}", s.Index);
                var styleName = string.Format("ta{0}", s.Index);

                file.automaticStyles.automaticStyles.Add(new ODContent.AutomaticStyles.Style()
                {
                    Name = styleName,
                    Family = "table",
                    MasterPageName = masterPageName,
                    tableProperties = new()
                });

                var table = new ODContent.Table()
                {
                    Name = s.SheetName,
                    StyleName = styleName
                };

                if (s.ColumnWidths.Any())
                {
                    table.tableColumns.Clear();
                    int colCount = s.Cells.Any() ? s.ColumnCount : 0;
                    int maxCol = Math.Max(colCount, s.ColumnWidths.Any() ? s.ColumnWidths.Keys.Max() : 0);

                    for (int c = 1; c <= maxCol; c++)
                    {
                        if (s.ColumnWidths.TryGetValue(c, out double width))
                        {
                            var colStyleName = $"co{s.Index}c{c}";
                            file.automaticStyles.automaticStyles.Add(new ODContent.AutomaticStyles.Style()
                            {
                                Name = colStyleName,
                                Family = "table-column",
                                tableColumnProperties = new() { ColumnWidth = $"{CharsToOdsCm(width)}cm" }
                            });
                            table.tableColumns.Add(new ODContent.Table.TableColumn() { StyleName = colStyleName, NumberColumnsRepeated = null });
                        }
                        else
                        {
                            table.tableColumns.Add(new ODContent.Table.TableColumn() { NumberColumnsRepeated = null });
                        }
                    }

                    int remaining = 16384 - maxCol;
                    if (remaining > 0)
                        table.tableColumns.Add(new ODContent.Table.TableColumn() { NumberColumnsRepeated = remaining.ToString() });
                }

                if (s.Cells.Any())
                {
                    int rowCount = s.RowCount;
                    int colCount = s.ColumnCount;

                    for (int r = 1; r <= rowCount; r++)
                    {
                        var tableRow = new ODContent.Table.TableRow();

                        for (int c = 1; c <= colCount; c++)
                        {
                            var cell = s.Cells.FirstOrDefault(x => x.Row == r && x.Column == c);

                            if (cell != null)
                            {
                                var cellStyleName = "ce1";
                                if (cell.ValueType == CellValueType.DateTime)
                                {
                                    var datePrefix = IsDateOnly(cell.Value) ? "do|" : "dt|";
                                    var visualKey = cell.Style != null ? GetStyleKey(cell.Style) : "";
                                    if (cellStyleMap.TryGetValue($"{datePrefix}{visualKey}", out var mapped))
                                        cellStyleName = mapped;
                                }
                                else if (cell.Style != null)
                                {
                                    var key = GetStyleKey(cell.Style);
                                    if (cellStyleMap.TryGetValue(key, out var mapped))
                                        cellStyleName = mapped;
                                }

                                var tableCell = new ODContent.Table.TableRow.TableCell()
                                {
                                    StyleName = cellStyleName,
                                    TextValue = cell.Value
                                };

                                if (cell.ValueType == CellValueType.Float || cell.ValueType == CellValueType.Int64)
                                {
                                    tableCell.ValueType = "float";
                                    tableCell.NumericValue = cell.Value;
                                }
                                else if (cell.ValueType == CellValueType.Boolean)
                                {
                                    tableCell.ValueType = "boolean";
                                    tableCell.BooleanValue = cell.Value.Equals("true", StringComparison.OrdinalIgnoreCase) ? "true" : "false";
                                }
                                else if (cell.ValueType == CellValueType.DateTime)
                                {
                                    tableCell.ValueType = "date";
                                    tableCell.DateValue = cell.Value;
                                }
                                else
                                {
                                    tableCell.ValueType = "string";
                                }

                                if (FormulaTranslator.Normalize(cell.Formula) is string formula)
                                {
                                    tableCell.Formula = FormulaTranslator.ToOpenFormula(formula);

                                    // Without a cached result, leave the cell untyped so the application calculates it on load
                                    if (string.IsNullOrEmpty(cell.Value))
                                    {
                                        tableCell.ValueType = null;
                                        tableCell.NumericValue = null;
                                        tableCell.BooleanValue = null;
                                        tableCell.DateValue = null;
                                        tableCell.TextValue = null;
                                    }
                                }

                                tableRow.Cells.Add(tableCell);
                            }
                            else
                            {
                                tableRow.Cells.Add(new ODContent.Table.TableRow.TableCell());
                            }
                        }

                        // Trailing empty columns filler
                        tableRow.Cells.Add(new ODContent.Table.TableRow.TableCell()
                        {
                            NumberColumnsRepeated = (16384 - colCount).ToString()
                        });

                        table.Rows.Add(tableRow);
                    }

                    // Trailing empty rows filler
                    table.Rows.Add(new ODContent.Table.TableRow()
                    {
                        NumberRowsRepeated = (1048576 - rowCount).ToString(),
                        Cells = new List<ODContent.Table.TableRow.TableCell>()
                        {
                            new ODContent.Table.TableRow.TableCell()
                            {
                                NumberColumnsRepeated = "16384"
                            }
                        }
                    });
                }
                else
                {
                    // Empty sheet — filler row
                    table.Rows.Add(new ODContent.Table.TableRow()
                    {
                        NumberRowsRepeated = "1048576",
                        Cells = new List<ODContent.Table.TableRow.TableCell>()
                        {
                            new ODContent.Table.TableRow.TableCell()
                            {
                                NumberColumnsRepeated = "16384"
                            }
                        }
                    });
                }

                file.body.spreadsheet.Tables.Add(table);
            }

            var dbRanges = new List<ODContent.Body.Spreadsheet.DatabaseRanges.DatabaseRange>();
            int dbIndex = 0;
            foreach (var s in workbook.Sheets)
            {
                if (s.AutoFilterRange is var (sr, sc, er, ec))
                {
                    var quotedName = s.SheetName.Contains(' ') ? $"'{s.SheetName}'" : s.SheetName;
                    var addr = $"{quotedName}.{OdsColumnLetter(sc)}{sr}:{quotedName}.{OdsColumnLetter(ec)}{er}";
                    dbRanges.Add(new ODContent.Body.Spreadsheet.DatabaseRanges.DatabaseRange
                    {
                        Name = $"__Anonymous_Sheet_DB__{dbIndex++}",
                        TargetRangeAddress = addr,
                        DisplayFilterButtons = "true"
                    });
                }
            }

            if (dbRanges.Any())
            {
                file.body.spreadsheet.databaseRanges = new ODContent.Body.Spreadsheet.DatabaseRanges();
                file.body.spreadsheet.databaseRanges.Ranges = dbRanges;
            }

            return (file, hasDateOnly, hasDateTime);
        }

        /// <summary>
        /// Generates meta.xml file found in the root of the ODS zip file
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private ODMeta GenerateMetaFile(oWorkbook workbook)
        {
            return new ODMeta()
            {
                meta = new ODMeta.Meta()
                {
                    CreationDate = workbook.CreationDate,
                    Creator = workbook.Creator,
                    Date = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ"),
                    Generator = workbook.Generator
                }
            };
        }

        private ODStyles GenerateStyleFile(oWorkbook workbook)
        {
            var file = new ODStyles();

            foreach (var s in workbook.Sheets)
            {
                var masterPageName = string.Format("mp{0}", s.Index);
                var pageLayoutName = string.Format("pm{0}", s.Index);

                file.automaticStyles.pageLayout.Add(new ODStyles.AutomaticStyles.PageLayout()
                {
                    Name = pageLayoutName
                });

                file.masterStyles.masterPage.Add(new ODStyles.MasterStyles.MasterPage()
                {
                    FooterLeftPageStyle = new ODStyles.MasterStyles.MasterPage.PageStyle()
                    {
                        Display = "false"
                    },
                    HeaderLeftPageStyle = new ODStyles.MasterStyles.MasterPage.PageStyle()
                    {
                        Display = "false"
                    },
                    Name = masterPageName,
                    PageLayoutName = pageLayoutName
                });
            }

            return file;
        }

        private const string OdsDateStyleName = "NDdate";
        private const string OdsDateTimeStyleName = "NDdatetime";

        private static Dictionary<string, string> BuildOdsCellStyles(oWorkbook workbook, ODContent file, out bool hasDateOnly, out bool hasDateTime)
        {
            var map = new Dictionary<string, string>();
            var nextIndex = 2;
            hasDateOnly = false;
            hasDateTime = false;

            foreach (var sheet in workbook.Sheets)
                foreach (var cell in sheet.Cells)
                {
                    var cs = cell.Style ?? new CellStyle();
                    var visualKey = cell.Style != null ? GetStyleKey(cs) : null;
                    string? datePrefix = null;

                    if (cell.ValueType == CellValueType.DateTime)
                    {
                        bool dateOnly = IsDateOnly(cell.Value);
                        datePrefix = dateOnly ? "do|" : "dt|";
                        if (dateOnly) hasDateOnly = true; else hasDateTime = true;
                    }

                    var mapKey = datePrefix != null
                        ? $"{datePrefix}{visualKey ?? ""}"
                        : visualKey;

                    if (mapKey == null) continue;
                    if (map.ContainsKey(mapKey)) continue;

                    var name = $"ce{nextIndex++}";
                    map[mapKey] = name;

                    var dataStyleName = datePrefix == "do|" ? OdsDateStyleName
                        : datePrefix == "dt|" ? OdsDateTimeStyleName
                        : "N0";

                    var style = new ODContent.AutomaticStyles.Style
                    {
                        Name = name,
                        Family = "table-cell",
                        ParentStyleName = "Default",
                        DataStyleName = dataStyleName
                    };

                    if (cs.Bold || cs.Italic || cs.Underline || cs.FontColor != null || cs.FontName != null || cs.FontSize != null)
                    {
                        style.textProperties = new ODContent.AutomaticStyles.Style.TextProperties();
                        if (cs.Bold) style.textProperties.FontWeight = "bold";
                        if (cs.Italic) style.textProperties.FontStyle = "italic";
                        if (cs.Underline)
                        {
                            style.textProperties.TextUnderlineStyle = "solid";
                            style.textProperties.TextUnderlineWidth = "auto";
                        }
                        if (cs.FontColor != null) style.textProperties.Color = cs.FontColor;
                        if (cs.FontName != null) style.textProperties.FontName = cs.FontName;
                        if (cs.FontSize != null) style.textProperties.FontSize = $"{cs.FontSize}pt";
                    }

                    if (cs.BackgroundColor != null || cs.BorderTop != null || cs.BorderBottom != null || cs.BorderLeft != null || cs.BorderRight != null || cs.WrapText)
                    {
                        style.tableCellProperties = new ODContent.AutomaticStyles.Style.TableCellStyleProperties();
                        if (cs.BackgroundColor != null) style.tableCellProperties.BackgroundColor = cs.BackgroundColor;
                        if (cs.BorderTop != null) style.tableCellProperties.BorderTop = FormatOdsBorder(cs.BorderTop);
                        if (cs.BorderBottom != null) style.tableCellProperties.BorderBottom = FormatOdsBorder(cs.BorderBottom);
                        if (cs.BorderLeft != null) style.tableCellProperties.BorderLeft = FormatOdsBorder(cs.BorderLeft);
                        if (cs.BorderRight != null) style.tableCellProperties.BorderRight = FormatOdsBorder(cs.BorderRight);
                        if (cs.WrapText) style.tableCellProperties.WrapOption = "wrap";
                    }

                    file.automaticStyles.automaticStyles.Add(style);
                }

            return map;
        }

        private static bool IsDateOnly(string value) =>
            DateTime.TryParse(value, out var dt) && dt.TimeOfDay == TimeSpan.Zero && !value.Contains('T');

        private static byte[] InjectOdsDateStyles(byte[] contentXml, bool hasDateOnly, bool hasDateTime)
        {
            if (!hasDateOnly && !hasDateTime) return contentXml;

            var xml = Encoding.UTF8.GetString(contentXml);
            var sb = new StringBuilder();

            if (hasDateOnly)
            {
                sb.Append($"<number:date-style style:name=\"{OdsDateStyleName}\">");
                sb.Append("<number:year number:style=\"long\"/>");
                sb.Append("<number:text>-</number:text>");
                sb.Append("<number:month number:style=\"long\"/>");
                sb.Append("<number:text>-</number:text>");
                sb.Append("<number:day number:style=\"long\"/>");
                sb.Append("</number:date-style>");
            }

            if (hasDateTime)
            {
                sb.Append($"<number:date-style style:name=\"{OdsDateTimeStyleName}\">");
                sb.Append("<number:year number:style=\"long\"/>");
                sb.Append("<number:text>-</number:text>");
                sb.Append("<number:month number:style=\"long\"/>");
                sb.Append("<number:text>-</number:text>");
                sb.Append("<number:day number:style=\"long\"/>");
                sb.Append("<number:text> </number:text>");
                sb.Append("<number:hours number:style=\"long\"/>");
                sb.Append("<number:text>:</number:text>");
                sb.Append("<number:minutes number:style=\"long\"/>");
                sb.Append("<number:text>:</number:text>");
                sb.Append("<number:seconds number:style=\"long\"/>");
                sb.Append("</number:date-style>");
            }

            xml = xml.Replace("<office:automatic-styles>", $"<office:automatic-styles>{sb}");
            return Encoding.UTF8.GetBytes(xml);
        }

        private static string GetStyleKey(CellStyle s) =>
            $"{s.Bold}|{s.Italic}|{s.Underline}|{s.FontColor}|{s.BackgroundColor}|{s.FontName}|{s.FontSize}|{s.WrapText}|{EdgeKey(s.BorderTop)}|{EdgeKey(s.BorderBottom)}|{EdgeKey(s.BorderLeft)}|{EdgeKey(s.BorderRight)}";

        private static string EdgeKey(CellBorder? b) =>
            b == null ? "" : $"{b.Style}:{b.Color}";

        private static double CharsToOdsCm(double chars) => Math.Round(chars * (1.69333333333333 / 8.43), 4);

        private static string OdsColumnLetter(int col)
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

        /// <summary>
        /// Returns a table's rows in document order, including rows nested inside
        /// table-header-rows, table-rows, and table-row-group containers (which may themselves be nested).
        /// </summary>
        private static IEnumerable<XElement> GetTableRows(XElement container, XNamespace tableNs)
        {
            foreach (var child in container.Elements())
            {
                if (child.Name == tableNs + "table-row")
                    yield return child;
                else if (child.Name == tableNs + "table-header-rows" || child.Name == tableNs + "table-rows" || child.Name == tableNs + "table-row-group")
                    foreach (var row in GetTableRows(child, tableNs))
                        yield return row;
            }
        }

        private static (string sheetName, int row, int col) ParseOdsCellAddress(string address)
        {
            var dotIdx = address.LastIndexOf('.');
            var sheetName = dotIdx >= 0 ? address[..dotIdx] : "";
            var cellRef = dotIdx >= 0 ? address[(dotIdx + 1)..] : address;

            int i = 0;
            while (i < cellRef.Length && cellRef[i] == '$') i++;
            int colStart = i;
            while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
            var colPart = cellRef[colStart..i];

            while (i < cellRef.Length && cellRef[i] == '$') i++;
            var rowPart = cellRef[i..];

            int col = 0;
            foreach (char c in colPart)
                col = col * 26 + (char.ToUpper(c) - 'A' + 1);

            return (sheetName, int.Parse(rowPart), col);
        }

        private static string FormatOdsBorder(CellBorder b)
        {
            if (b.Style == BorderStyle.None) return "none";
            var width = b.Style switch
            {
                BorderStyle.Thin => "0.75pt",
                BorderStyle.Medium => "1.5pt",
                BorderStyle.Thick => "2.5pt",
                _ => "0.75pt"
            };
            var color = b.Color ?? "#000000";
            return $"{width} solid {color}";
        }

        private static byte[] BuildSettingsFile(oWorkbook workbook)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.Append("<office:document-settings xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:config=\"urn:oasis:names:tc:opendocument:xmlns:config:1.0\" office:version=\"1.3\">");
            sb.Append("<office:settings>");
            sb.Append("<config:config-item-set config:name=\"ooo:view-settings\">");
            sb.Append("<config:config-item-map-indexed config:name=\"Views\">");
            sb.Append("<config:config-item-map-entry>");
            sb.Append("<config:config-item-map-named config:name=\"Tables\">");

            foreach (var sheet in workbook.Sheets)
            {
                if (sheet.FreezeRows <= 0 && sheet.FreezeColumns <= 0) continue;

                sb.Append($"<config:config-item-map-entry config:name=\"{SecurityElement.Escape(sheet.SheetName)}\">");
                sb.Append($"<config:config-item config:name=\"HorizontalSplitMode\" config:type=\"short\">2</config:config-item>");
                sb.Append($"<config:config-item config:name=\"VerticalSplitMode\" config:type=\"short\">2</config:config-item>");
                sb.Append($"<config:config-item config:name=\"HorizontalSplitPosition\" config:type=\"int\">{sheet.FreezeColumns}</config:config-item>");
                sb.Append($"<config:config-item config:name=\"VerticalSplitPosition\" config:type=\"int\">{sheet.FreezeRows}</config:config-item>");
                sb.Append($"<config:config-item config:name=\"PositionRight\" config:type=\"int\">{sheet.FreezeColumns}</config:config-item>");
                sb.Append($"<config:config-item config:name=\"PositionBottom\" config:type=\"int\">{sheet.FreezeRows}</config:config-item>");
                sb.Append("</config:config-item-map-entry>");
            }

            sb.Append("</config:config-item-map-named>");
            sb.Append("</config:config-item-map-entry>");
            sb.Append("</config:config-item-map-indexed>");
            sb.Append("</config:config-item-set>");
            sb.Append("</office:settings>");
            sb.Append("</office:document-settings>");

            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~GenerateOdsFileService()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
