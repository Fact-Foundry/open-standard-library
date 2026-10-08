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
                var (content, hasDateOnly, hasDateTime, customDataStyles) = GenerateContentFile(workbook);
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
                        Content = InjectOdsDataStyles(await XmlService.ConvertToXmlAsync(content), hasDateOnly, hasDateTime, customDataStyles)
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
                doc = await Task.Run(() => XDocument.Load(contentStream, LoadOptions.PreserveWhitespace));

            XNamespace tableNs  = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
            XNamespace officeNs = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XNamespace textNs   = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

            var cellStyles = await ReadCellStylesAsync(doc, archive);
            var columnWidths = ReadColumnWidths(doc);

            foreach (var table in doc.Descendants(tableNs + "table"))
            {
                var sheetName = table.Attribute(tableNs + "name")?.Value ?? "Sheet";
                var sheet = workbook.AddSheet(sheetName);

                // Walk the column declarations (which may be nested in header/group wrappers) in order, collecting
                // widths and each column's default cell style, which applications use instead of styling every cell.
                // A trailing repeated declaration is filler for the unused columns and is skipped.
                var columnDefaultStyles = new Dictionary<int, string>();
                var columns = table.Descendants(tableNs + "table-column").ToList();
                int columnIndex = 0;
                for (int i = 0; i < columns.Count; i++)
                {
                    var column = columns[i];
                    int repeated = int.TryParse(column.Attribute(tableNs + "number-columns-repeated")?.Value, out int cr) ? cr : 1;
                    bool trailingFiller = i == columns.Count - 1 && repeated > 1;
                    int last = Math.Min(columnIndex + repeated, columnIndex + MaxImportedWidthColumns);

                    var columnStyle = column.Attribute(tableNs + "style-name")?.Value;
                    if (!trailingFiller && columnStyle != null && columnWidths.TryGetValue(columnStyle, out double width))
                        for (int c = columnIndex + 1; c <= last; c++)
                            sheet.SetColumnWidth(c, width);

                    var defaultCellStyle = column.Attribute(tableNs + "default-cell-style-name")?.Value;
                    if (!trailingFiller && defaultCellStyle != null && cellStyles.ContainsKey(defaultCellStyle))
                        for (int c = columnIndex + 1; c <= last; c++)
                            columnDefaultStyles[c] = defaultCellStyle;

                    columnIndex += repeated;
                }

                int rowIndex = 0;

                foreach (var tableRow in GetTableRows(table, tableNs))
                {
                    int rowsRepeated = int.TryParse(tableRow.Attribute(tableNs + "number-rows-repeated")?.Value, out int rr) ? rr : 1;

                    var rowData = new List<(int col, string value, CellValueType type, string? formula, string? styleName)>();
                    int colIndex = 0;

                    foreach (var cell in tableRow.Elements().Where(e => e.Name == tableNs + "table-cell" || e.Name == tableNs + "covered-table-cell"))
                    {
                        int colsRepeated = int.TryParse(cell.Attribute(tableNs + "number-columns-repeated")?.Value, out int cr) ? cr : 1;

                        // Covered cells are hidden by a merge; they occupy column positions but hold no visible data
                        if (cell.Name == tableNs + "covered-table-cell")
                        {
                            colIndex += colsRepeated;
                            continue;
                        }

                        var valueType  = cell.Attribute(officeNs + "value-type")?.Value;
                        var textValue  = GetCellText(cell, textNs);
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
                            else if (valueType is "float" or "percentage" or "currency")
                            {
                                // office:value holds the actual number; the paragraph text is only the formatted display (e.g. "1,234.50", "25%", "$9.99")
                                cellType = CellValueType.Float;
                                cellValue = numericValue ?? textValue ?? "";
                            }
                            else
                            {
                                cellType = CellValueType.String;
                                cellValue = textValue ?? numericValue ?? "";
                            }

                            var styleName = cell.Attribute(tableNs + "style-name")?.Value;
                            for (int i = 0; i < colsRepeated; i++)
                            {
                                colIndex++;
                                rowData.Add((colIndex, cellValue, cellType, formula, styleName));
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
                        foreach (var (col, value, type, formula, styleName) in rowData)
                        {
                            var oCell = sheet.AddCell(rowIndex, col, value);
                            oCell.ValueType = type;
                            oCell.Formula = formula;
                            var effectiveStyleName = styleName ?? (columnDefaultStyles.TryGetValue(col, out var columnDefault) ? columnDefault : null);

                            if (effectiveStyleName != null && cellStyles.TryGetValue(effectiveStyleName, out var cellStyle))
                            {
                                var style = cellStyle.Clone();
                                if (style.NumberFormat != null && IsDefaultFormat(style.NumberFormat, type))
                                    style.NumberFormat = null;
                                if (!style.IsDefault)
                                    oCell.Style = style;
                            }
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

        private (ODContent content, bool hasDateOnly, bool hasDateTime, List<string> customDataStyles) GenerateContentFile(oWorkbook workbook)
        {
            var file = new ODContent();

            var cellStyleMap = BuildOdsCellStyles(workbook, file, out bool hasDateOnly, out bool hasDateTime, out var customDataStyles);

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
                            var cell = s.GetCell(r, c);

                            if (cell != null)
                            {
                                var cellStyleName = "ce1";
                                if (CellStyleMapKey(cell) is string mapKey && cellStyleMap.TryGetValue(mapKey, out var mapped))
                                    cellStyleName = mapped;

                                var tableCell = new ODContent.Table.TableRow.TableCell()
                                {
                                    StyleName = cellStyleName,
                                    TextValue = cell.Value
                                };

                                if (cell.ValueType == CellValueType.Float || cell.ValueType == CellValueType.Int64)
                                {
                                    // Percentage and currency formats have their own ODS value types
                                    var formatCode = CustomFormatCode(cell);
                                    tableCell.ValueType = (formatCode != null ? NumberFormatTranslator.OdsValueType(formatCode) : null) ?? "float";
                                    tableCell.NumericValue = cell.Value;
                                    if (tableCell.ValueType == "currency")
                                        tableCell.Currency = NumberFormatTranslator.CurrencyCode(NumberFormatTranslator.Parse(formatCode!).CurrencySymbol);
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

            return (file, hasDateOnly, hasDateTime, customDataStyles);
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

        private static Dictionary<string, string> BuildOdsCellStyles(oWorkbook workbook, ODContent file, out bool hasDateOnly, out bool hasDateTime, out List<string> customDataStyles)
        {
            var map = new Dictionary<string, string>();
            var nextIndex = 2;
            hasDateOnly = false;
            hasDateTime = false;

            // One data style per distinct format code, named N100, N101, ...
            var dataStyleNames = new Dictionary<string, string>();
            customDataStyles = new List<string>();

            foreach (var sheet in workbook.Sheets)
                foreach (var cell in sheet.Cells)
                {
                    var cs = cell.Style ?? new CellStyle();
                    var mapKey = CellStyleMapKey(cell);
                    if (mapKey == null) continue;

                    var formatCode = CustomFormatCode(cell);
                    if (formatCode == null && cell.ValueType == CellValueType.DateTime)
                    {
                        if (IsDateOnly(cell.Value)) hasDateOnly = true; else hasDateTime = true;
                    }

                    if (map.ContainsKey(mapKey)) continue;

                    var name = $"ce{nextIndex++}";
                    map[mapKey] = name;

                    string dataStyleName;
                    if (formatCode != null)
                    {
                        if (!dataStyleNames.TryGetValue(formatCode, out dataStyleName!))
                        {
                            dataStyleName = $"N{100 + dataStyleNames.Count}";
                            dataStyleNames[formatCode] = dataStyleName;
                            customDataStyles.Add(NumberFormatTranslator.ToOdsDataStyle(formatCode, dataStyleName)!);
                        }
                    }
                    else
                        dataStyleName = mapKey.StartsWith("do|") ? OdsDateStyleName
                            : mapKey.StartsWith("dt|") ? OdsDateTimeStyleName
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

        /// <summary>
        /// Inserts the data style elements (default date styles and custom number formats) into office:automatic-styles.
        /// They are written as raw XML because the XmlSerializer content model has no classes for number:*-style elements.
        /// </summary>
        private static byte[] InjectOdsDataStyles(byte[] contentXml, bool hasDateOnly, bool hasDateTime, List<string> customDataStyles)
        {
            if (!hasDateOnly && !hasDateTime && customDataStyles.Count == 0) return contentXml;

            var xml = Encoding.UTF8.GetString(contentXml);
            var sb = new StringBuilder();

            foreach (var dataStyle in customDataStyles)
                sb.Append(dataStyle);

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
            $"{s.Bold}|{s.Italic}|{s.Underline}|{s.FontColor}|{s.BackgroundColor}|{s.FontName}|{s.FontSize}|{s.WrapText}|{EdgeKey(s.BorderTop)}|{EdgeKey(s.BorderBottom)}|{EdgeKey(s.BorderLeft)}|{EdgeKey(s.BorderRight)}|{s.NumberFormat}";

        /// <summary>
        /// A cell's NumberFormat when it is set and can be expressed as an ODS data style; otherwise null and the defaults apply.
        /// </summary>
        private static string? CustomFormatCode(oCell cell) =>
            !string.IsNullOrEmpty(cell.Style?.NumberFormat) && NumberFormatTranslator.ToOdsDataStyle(cell.Style.NumberFormat, "x") != null
                ? cell.Style.NumberFormat
                : null;

        /// <summary>
        /// Key into the cell style map: the visual style key, prefixed by the default date style kind when no custom format applies.
        /// </summary>
        private static string? CellStyleMapKey(oCell cell)
        {
            var visualKey = cell.Style != null ? GetStyleKey(cell.Style) : null;
            if (cell.ValueType == CellValueType.DateTime && CustomFormatCode(cell) == null)
                return (IsDateOnly(cell.Value) ? "do|" : "dt|") + (visualKey ?? "");
            return visualKey;
        }

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

        // A column declaration can repeat for the whole sheet (LibreOffice writes 1024 or more); cap how many widths one sets
        private const int MaxImportedWidthColumns = 1024;

        /// <summary>
        /// Maps each table-column style name to its width in characters. The library's own default column width is skipped
        /// so a round-trip of a sheet with no explicit widths leaves ColumnWidths empty.
        /// </summary>
        private static Dictionary<string, double> ReadColumnWidths(XDocument contentDoc)
        {
            XNamespace styleNs = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            var result = new Dictionary<string, double>();

            foreach (var style in contentDoc.Descendants(styleNs + "style").Where(s => (string?)s.Attribute(styleNs + "family") == "table-column"))
            {
                var name = (string?)style.Attribute(styleNs + "name");
                var widthText = (string?)style.Element(styleNs + "table-column-properties")?.Attribute(styleNs + "column-width");
                if (name == null || widthText == null) continue;

                var cm = ParseLengthCm(widthText);
                if (cm == null) continue;

                var chars = Math.Round(cm.Value * (8.43 / 1.69333333333333), 2);
                if (Math.Abs(chars - 8.43) < 0.01) continue; // the library's default column

                result[name] = chars;
            }
            return result;
        }

        /// <summary>
        /// Parses an ODF length ("1.69cm", "0.5in", "12pt", "20mm") to centimeters.
        /// </summary>
        private static double? ParseLengthCm(string text)
        {
            var units = new (string Unit, double ToCm)[] { ("cm", 1), ("mm", 0.1), ("in", 2.54), ("pt", 2.54 / 72), ("pc", 2.54 / 6) };
            foreach (var (unit, toCm) in units)
                if (text.EndsWith(unit, StringComparison.Ordinal)
                    && double.TryParse(text[..^unit.Length], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value))
                    return value * toCm;
            return null;
        }

        /// <summary>
        /// Builds a CellStyle for each table-cell automatic style from its text properties, cell properties, and data style.
        /// Styles that resolve to the default are omitted. Data styles may live in content.xml or styles.xml.
        /// </summary>
        private static async Task<Dictionary<string, CellStyle>> ReadCellStylesAsync(XDocument contentDoc, ZipArchive archive)
        {
            XNamespace styleNs = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XNamespace numberNs = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
            XNamespace foNs = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

            var dataStyles = new Dictionary<string, XElement>();
            void CollectDataStyles(XDocument document)
            {
                foreach (var ds in document.Descendants().Where(e => e.Name.Namespace == numberNs && e.Name.LocalName.EndsWith("-style", StringComparison.Ordinal)))
                    if ((string?)ds.Attribute(styleNs + "name") is string name)
                        dataStyles[name] = ds;
            }

            CollectDataStyles(contentDoc);
            var stylesEntry = archive.GetEntry("styles.xml");
            if (stylesEntry != null)
            {
                using var stylesStream = stylesEntry.Open();
                CollectDataStyles(await Task.Run(() => XDocument.Load(stylesStream, LoadOptions.PreserveWhitespace)));
            }

            var result = new Dictionary<string, CellStyle>();
            foreach (var cellStyle in contentDoc.Descendants(styleNs + "style").Where(s => (string?)s.Attribute(styleNs + "family") == "table-cell"))
            {
                var name = (string?)cellStyle.Attribute(styleNs + "name");
                if (name == null) continue;

                var style = new CellStyle();

                var text = cellStyle.Element(styleNs + "text-properties");
                if (text != null)
                {
                    style.Bold = (string?)text.Attribute(foNs + "font-weight") == "bold";
                    style.Italic = (string?)text.Attribute(foNs + "font-style") == "italic";
                    style.Underline = (string?)text.Attribute(styleNs + "text-underline-style") is string u && u != "none";
                    style.FontColor = NormalizeColor((string?)text.Attribute(foNs + "color"));
                    style.FontName = (string?)text.Attribute(styleNs + "font-name");
                    var size = (string?)text.Attribute(foNs + "font-size");
                    if (size != null && size.EndsWith("pt", StringComparison.Ordinal)
                        && double.TryParse(size[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double pt))
                        style.FontSize = pt;
                }

                var cellProps = cellStyle.Element(styleNs + "table-cell-properties");
                if (cellProps != null)
                {
                    style.BackgroundColor = NormalizeColor((string?)cellProps.Attribute(foNs + "background-color"));
                    style.WrapText = (string?)cellProps.Attribute(foNs + "wrap-option") == "wrap";

                    // fo:border applies to all four sides unless a side-specific attribute overrides it
                    var all = ParseOdsBorder((string?)cellProps.Attribute(foNs + "border"));
                    style.BorderTop = ParseOdsBorder((string?)cellProps.Attribute(foNs + "border-top")) ?? all?.Clone();
                    style.BorderBottom = ParseOdsBorder((string?)cellProps.Attribute(foNs + "border-bottom")) ?? all?.Clone();
                    style.BorderLeft = ParseOdsBorder((string?)cellProps.Attribute(foNs + "border-left")) ?? all?.Clone();
                    style.BorderRight = ParseOdsBorder((string?)cellProps.Attribute(foNs + "border-right")) ?? all?.Clone();
                }

                var dataStyleName = (string?)cellStyle.Attribute(styleNs + "data-style-name");
                if (dataStyleName != null && dataStyles.TryGetValue(dataStyleName, out var dataStyle)
                    && NumberFormatTranslator.FromOdsDataStyle(dataStyle) is string code)
                    style.NumberFormat = code;

                if (!style.IsDefault)
                    result[name] = style;
            }
            return result;
        }

        /// <summary>
        /// Returns a "#RRGGBB" color, or null for "transparent" and anything that isn't a hex color.
        /// </summary>
        private static string? NormalizeColor(string? color) =>
            color != null && color.StartsWith('#') && color.Length == 7 ? color.ToUpperInvariant() : null;

        /// <summary>
        /// Parses an ODF border such as "0.75pt solid #000000". Returns null for "none" or an unparseable value.
        /// </summary>
        private static CellBorder? ParseOdsBorder(string? border)
        {
            if (string.IsNullOrEmpty(border) || border == "none") return null;

            var parts = border.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var widthCm = parts.Select(ParseLengthCm).FirstOrDefault(w => w != null);
            var widthPt = (widthCm ?? 0) / 2.54 * 72;
            var style = widthPt switch
            {
                <= 1.0 => BorderStyle.Thin,
                <= 2.0 => BorderStyle.Medium,
                _ => BorderStyle.Thick
            };

            return new CellBorder { Style = style, Color = parts.Select(NormalizeColor).FirstOrDefault(c => c != null) };
        }

        /// <summary>
        /// The library's own default date formats aren't reported as a NumberFormat on import, so a round-trip leaves Style null.
        /// </summary>
        private static bool IsDefaultFormat(string code, CellValueType valueType) =>
            valueType == CellValueType.DateTime && code is NumberFormatTranslator.DefaultDateCode or NumberFormatTranslator.DefaultDateTimeCode;

        /// <summary>
        /// Returns a cell's text content, or null if it has no paragraphs. Paragraphs are joined with line feeds,
        /// and the ODF whitespace elements (text:s, text:tab, text:line-break) are expanded to the characters they represent.
        /// </summary>
        private static string? GetCellText(XElement cell, XNamespace textNs)
        {
            var paragraphs = cell.Elements(textNs + "p").ToList();
            if (paragraphs.Count == 0)
                return null;

            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (i > 0)
                    sb.Append('\n');
                AppendText(paragraphs[i], textNs, sb);
            }
            return sb.ToString();
        }

        private static void AppendText(XElement element, XNamespace textNs, System.Text.StringBuilder sb)
        {
            foreach (var node in element.Nodes())
            {
                if (node is XText text)
                    sb.Append(text.Value);
                else if (node is XElement child)
                {
                    if (child.Name == textNs + "s")
                        sb.Append(' ', int.TryParse(child.Attribute(textNs + "c")?.Value, out int count) ? count : 1);
                    else if (child.Name == textNs + "tab")
                        sb.Append('\t');
                    else if (child.Name == textNs + "line-break")
                        sb.Append('\n');
                    else
                        AppendText(child, textNs, sb); // text:span, text:a, etc.
                }
            }
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
