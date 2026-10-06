using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Structural validator for OpenDocument spreadsheets (ODF 1.2+). Checks the problems that make LibreOffice refuse a file
    /// or prompt to repair it: the mimetype entry, the manifest, document roots, table structure, and cell values.
    /// </summary>
    internal sealed class OdsValidator
    {
        private const string Prefix = "ODS";
        private const int MaxColumns = 16_384;
        private const int MaxRows = 1_048_576;

        private static readonly XNamespace OfficeNs = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        private static readonly XNamespace TableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        private static readonly XNamespace StyleNs = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        private static readonly XNamespace NumberNs = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
        private static readonly XNamespace ManifestNs = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        private static readonly XNamespace TextNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private static readonly XNamespace CalcExtNs = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";

        private static readonly HashSet<string> PropertiesElements = new(StringComparer.Ordinal)
        {
            "table-properties", "table-column-properties", "table-row-properties", "table-cell-properties", "text-properties",
            "paragraph-properties", "graphic-properties", "page-layout-properties", "header-footer-properties", "chart-properties",
            "drawing-page-properties", "ruby-properties", "section-properties", "list-level-properties"
        };

        private static readonly Dictionary<string, string> RootElements = new(StringComparer.Ordinal)
        {
            ["content.xml"] = "document-content",
            ["styles.xml"] = "document-styles",
            ["meta.xml"] = "document-meta",
            ["settings.xml"] = "document-settings"
        };

        // Containers that may hold table rows directly.
        private static readonly HashSet<string> RowContainers = new(StringComparer.Ordinal)
        {
            "table", "table-row-group", "table-header-rows", "table-rows"
        };

        private readonly ZipPackage _package;
        private readonly IssueCollector _issues;
        private readonly byte[] _file;
        private readonly HashSet<string> _styleNames = new(StringComparer.Ordinal);
        private readonly HashSet<string> _dataStyleNames = new(StringComparer.Ordinal);
        private readonly HashSet<string> _tableNames = new(StringComparer.OrdinalIgnoreCase);

        private OdsValidator(byte[] file, ZipPackage package, IssueCollector issues)
        {
            _file = file;
            _package = package;
            _issues = issues;
        }

        internal static void Validate(byte[] file, IssueCollector issues, ValidationOptions options)
        {
            if (!FileSniffer.IsZip(file))
            {
                issues.Error($"{Prefix}_NOT_ZIP",
                    $"An ODS file must be a ZIP archive, but this file is {FileSniffer.Describe(file)}. ODS files can't be written as plain text.");
                return;
            }

            using var package = ZipPackage.Open(file, issues, options, Prefix);
            if (package == null)
                return;

            new OdsValidator(file, package, issues).Run();
        }

        private void Run()
        {
            if (_package.Exists("[Content_Types].xml") && !_package.Exists("mimetype"))
            {
                _issues.Error($"{Prefix}_IS_XLSX", "This file is an Office Open XML workbook (XLSX), not ODS. Validate it as XLSX or save it with an .xlsx extension.");
                return;
            }

            if (!ValidateMimetype())
                return;

            foreach (var part in _package.PartNames.Where(n => n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
                _package.LoadXml(part);

            ValidateManifest();

            foreach (var (part, rootName) in RootElements)
                ValidateRoot(part, rootName, required: part == "content.xml");

            CollectStyleNames("styles.xml");
            CollectStyleNames("content.xml");

            ValidateContent();
        }

        // ---------------------------------------------------------------- Packaging

        private bool ValidateMimetype()
        {
            if (!_package.Exists("mimetype"))
            {
                _issues.Error($"{Prefix}_MIMETYPE_MISSING",
                    $"The archive has no \"mimetype\" entry. An ODS file must start with an uncompressed entry named \"mimetype\" containing \"{FileSniffer.OdsMimeType}\".");
                return true;
            }

            var bytes = _package.ReadBytes("mimetype");
            if (bytes == null)
                return true;

            var value = Encoding.ASCII.GetString(bytes);
            if (!IsSpreadsheetMimeType(value))
            {
                if (IsSpreadsheetMimeType(value.Trim()))
                    _issues.Error($"{Prefix}_MIMETYPE_WHITESPACE", "The mimetype entry contains whitespace or a newline. It must contain exactly the MIME type with no trailing characters.", "mimetype");
                else if (value.StartsWith("application/vnd.oasis.opendocument.", StringComparison.Ordinal))
                {
                    _issues.Error($"{Prefix}_NOT_SPREADSHEET", $"This is an OpenDocument file of type \"{value.Trim()}\", not a spreadsheet.", "mimetype");
                    return false;
                }
                else
                    _issues.Error($"{Prefix}_MIMETYPE_WRONG", $"The mimetype entry contains \"{XlsxValidator.Truncate(value.Trim())}\" but must contain \"{FileSniffer.OdsMimeType}\".", "mimetype");
            }

            // The first local file header must be the mimetype entry, stored (method 0) with no extra field.
            // ZipArchive does not expose compression method, so read the header directly.
            if (_file.Length >= 30)
            {
                var method = BinaryPrimitives.ReadUInt16LittleEndian(_file.AsSpan(8));
                var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(_file.AsSpan(26));
                var extraLength = BinaryPrimitives.ReadUInt16LittleEndian(_file.AsSpan(28));
                var firstName = _file.Length >= 30 + nameLength ? Encoding.UTF8.GetString(_file, 30, nameLength) : "";

                if (firstName != "mimetype")
                    _issues.Warning($"{Prefix}_MIMETYPE_NOT_FIRST",
                        $"The first entry in the archive is \"{firstName}\". The ODF specification requires \"mimetype\" to be the first entry so tools can identify the file type.", "mimetype");
                else
                {
                    if (method != 0)
                        _issues.Warning($"{Prefix}_MIMETYPE_COMPRESSED", "The mimetype entry is compressed. The ODF specification requires it to be stored without compression.", "mimetype");
                    if (extraLength != 0)
                        _issues.Warning($"{Prefix}_MIMETYPE_EXTRA_FIELD", "The mimetype entry's ZIP header has an extra field. The ODF specification says it should not.", "mimetype");
                }
            }

            return true;
        }

        private void ValidateManifest()
        {
            const string part = "META-INF/manifest.xml";
            if (!_package.Exists(part))
            {
                _issues.Error($"{Prefix}_MANIFEST_MISSING", "The archive has no META-INF/manifest.xml. It is required and must list every file in the package.");
                return;
            }

            var doc = _package.LoadXml(part);
            if (doc?.Root == null)
                return;

            if (doc.Root.Name != ManifestNs + "manifest")
            {
                _issues.Error($"{Prefix}_MANIFEST_ROOT",
                    $"The root element must be <manifest:manifest xmlns:manifest=\"{ManifestNs.NamespaceName}\"> but is <{doc.Root.Name.LocalName}> in namespace \"{doc.Root.Name.NamespaceName}\".", part);
                return;
            }

            var listed = new HashSet<string>(StringComparer.Ordinal);
            var hasRootEntry = false;

            foreach (var entry in doc.Root.Elements(ManifestNs + "file-entry"))
            {
                var path = (string?)entry.Attribute(ManifestNs + "full-path");
                var mediaType = (string?)entry.Attribute(ManifestNs + "media-type");

                if (path == null || mediaType == null)
                {
                    _issues.Error($"{Prefix}_MANIFEST_ENTRY_INVALID",
                        "<manifest:file-entry> requires manifest:full-path and manifest:media-type attributes.", part, ZipPackage.LineOf(entry));
                    continue;
                }

                if (path == "/")
                {
                    hasRootEntry = true;
                    if (!IsSpreadsheetMimeType(mediaType))
                        _issues.Error($"{Prefix}_MANIFEST_ROOT_MEDIA_TYPE",
                            $"The root entry (full-path=\"/\") has media-type \"{mediaType}\" but must be \"{FileSniffer.OdsMimeType}\".", part, ZipPackage.LineOf(entry));
                    continue;
                }

                if (!listed.Add(path))
                    _issues.Warning($"{Prefix}_MANIFEST_DUPLICATE", $"\"{path}\" is listed more than once.", part, ZipPackage.LineOf(entry));

                if (!path.EndsWith('/') && !_package.Exists(path))
                    _issues.Warning($"{Prefix}_MANIFEST_FILE_MISSING",
                        $"The manifest lists \"{path}\" but the archive has no such entry. Remove the entry or add the file.", part, ZipPackage.LineOf(entry));
            }

            if (!hasRootEntry)
                _issues.Error($"{Prefix}_MANIFEST_NO_ROOT_ENTRY",
                    $"The manifest has no entry for the document itself. Add <manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{FileSniffer.OdsMimeType}\"/>.", part);

            foreach (var name in _package.PartNames)
            {
                if (name == "mimetype" || name.StartsWith("META-INF/", StringComparison.Ordinal))
                    continue;
                if (!listed.Contains(name))
                    _issues.Error($"{Prefix}_MANIFEST_FILE_UNLISTED",
                        $"\"{name}\" is in the archive but not listed in META-INF/manifest.xml. LibreOffice treats unlisted files as corruption.", part);
            }
        }

        private void ValidateRoot(string part, string rootName, bool required)
        {
            if (!_package.Exists(part))
            {
                if (required)
                    _issues.Error($"{Prefix}_CONTENT_MISSING", "The archive has no content.xml, which holds the spreadsheet data.");
                return;
            }

            var doc = _package.LoadXml(part);
            if (doc?.Root == null)
                return;

            if (doc.Root.Name != OfficeNs + rootName)
            {
                _issues.Error($"{Prefix}_ROOT_ELEMENT",
                    $"The root element must be <office:{rootName}> in namespace \"{OfficeNs.NamespaceName}\" but is <{doc.Root.Name.LocalName}> in namespace \"{doc.Root.Name.NamespaceName}\".",
                    part);
                return;
            }

            if (doc.Root.Attribute(OfficeNs + "version") == null)
                _issues.Warning($"{Prefix}_VERSION_MISSING", $"<office:{rootName}> has no office:version attribute (for example office:version=\"1.3\").", part);

            // LibreOffice legitimately writes some properties in its own extension namespaces (e.g. loext:graphic-properties),
            // so only flag properties placed in a standard ODF namespace other than style.
            foreach (var element in doc.Root.Descendants().Where(e => PropertiesElements.Contains(e.Name.LocalName)
                && e.Name.Namespace != StyleNs && e.Name.NamespaceName.StartsWith("urn:oasis:names:tc:opendocument:", StringComparison.Ordinal)))
                _issues.Warning($"{Prefix}_PROPERTIES_NAMESPACE",
                    $"<{element.Name.LocalName}> is in namespace \"{element.Name.NamespaceName}\" but formatting properties must be in the style namespace (style:{element.Name.LocalName}). Applications ignore the formatting it contains.",
                    part, ZipPackage.LineOf(element));
        }

        private void CollectStyleNames(string part)
        {
            if (!_package.Exists(part) || _package.LoadXml(part)?.Root is not XElement root || root.Name.Namespace != OfficeNs)
                return;

            foreach (var style in root.Descendants(StyleNs + "style"))
                if ((string?)style.Attribute(StyleNs + "name") is string name)
                    _styleNames.Add(name);

            foreach (var dataStyle in root.Descendants().Where(e => e.Name.Namespace == NumberNs && e.Name.LocalName.EndsWith("-style", StringComparison.Ordinal)))
                if ((string?)dataStyle.Attribute(StyleNs + "name") is string name)
                    _dataStyleNames.Add(name);
        }

        // ---------------------------------------------------------------- Content

        private void ValidateContent()
        {
            const string part = "content.xml";
            if (_package.LoadXml(part)?.Root is not XElement root || root.Name != OfficeNs + "document-content")
                return;

            var body = root.Element(OfficeNs + "body");
            var spreadsheet = body?.Element(OfficeNs + "spreadsheet");
            if (spreadsheet == null)
            {
                var other = body?.Elements().FirstOrDefault()?.Name.LocalName;
                _issues.Error($"{Prefix}_NO_SPREADSHEET_BODY", other != null
                    ? $"<office:body> contains <office:{other}> instead of <office:spreadsheet>, so this is not a spreadsheet document."
                    : "content.xml must contain <office:body><office:spreadsheet>...</office:spreadsheet></office:body>.", part);
                return;
            }

            foreach (var style in root.Descendants(StyleNs + "style"))
            {
                var dataStyle = style.Attribute(StyleNs + "data-style-name");
                if (dataStyle != null && !_dataStyleNames.Contains(dataStyle.Value))
                    _issues.Warning($"{Prefix}_DATA_STYLE_MISSING",
                        $"Style \"{(string?)style.Attribute(StyleNs + "name")}\" refers to data style \"{dataStyle.Value}\", which is not defined.", part, ZipPackage.LineOf(dataStyle));
            }

            var tables = spreadsheet.Elements(TableNs + "table").ToList();
            if (tables.Count == 0)
            {
                _issues.Error($"{Prefix}_NO_TABLES", "<office:spreadsheet> contains no <table:table> elements. A spreadsheet must contain at least one sheet.", part);
                return;
            }

            foreach (var table in tables)
                if ((string?)table.Attribute(TableNs + "name") is string tableName)
                    _tableNames.Add(tableName);

            var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var table in tables)
            {
                var name = (string?)table.Attribute(TableNs + "name");
                if (string.IsNullOrEmpty(name))
                    _issues.Error($"{Prefix}_TABLE_NAME_MISSING", "A <table:table> has no table:name attribute.", part, ZipPackage.LineOf(table));
                else
                {
                    if (ValidateTableName(name) is string problem)
                        _issues.Error($"{Prefix}_TABLE_NAME_INVALID", $"Sheet name \"{name}\" {problem}", part, ZipPackage.LineOf(table));
                    if (!tableNames.Add(name))
                        _issues.Error($"{Prefix}_TABLE_NAME_DUPLICATE", $"Sheet name \"{name}\" is used more than once. Sheet names must be unique.", part, ZipPackage.LineOf(table));
                }

                ValidateTable(table, name ?? "?", part);
            }

            foreach (var range in spreadsheet.Descendants(TableNs + "database-range"))
            {
                var address = (string?)range.Attribute(TableNs + "target-range-address");
                if (address == null || !IsValidRangeAddress(address))
                    _issues.Warning($"{Prefix}_RANGE_ADDRESS_INVALID",
                        $"table:target-range-address=\"{address}\" is not a valid range such as Sheet1.A1:Sheet1.C10.", part, ZipPackage.LineOf(range));
            }
        }

        private static bool IsSpreadsheetMimeType(string value) =>
            value is FileSniffer.OdsMimeType or FileSniffer.OdsTemplateMimeType;

        private static string? ValidateTableName(string name)
        {
            var invalid = name.IndexOfAny(['[', ']', '*', '?', ':', '/', '\\']);
            if (invalid >= 0)
                return $"contains '{name[invalid]}'. Sheet names can't contain [ ] * ? : / or \\.";
            if (name.StartsWith('\'') || name.EndsWith('\''))
                return "starts or ends with an apostrophe, which is not allowed.";
            return null;
        }

        private void ValidateTable(XElement table, string tableName, string part)
        {
            CheckStyleReference(table, TableNs + "style-name", part);

            var totalRows = 0L;

            foreach (var element in table.Descendants())
            {
                if (element.Name.Namespace != TableNs)
                    continue;

                switch (element.Name.LocalName)
                {
                    case "table-column":
                        CheckStyleReference(element, TableNs + "style-name", part);
                        CheckStyleReference(element, TableNs + "default-cell-style-name", part);
                        CheckPositiveInteger(element, "number-columns-repeated", part);
                        break;

                    case "table-row":
                        if (!RowContainers.Contains(element.Parent!.Name.LocalName) || element.Parent.Name.Namespace != TableNs)
                            _issues.Error($"{Prefix}_ROW_MISPLACED",
                                $"<table:table-row> is inside <{element.Parent.Name.LocalName}>. Rows must be children of <table:table> or a row group.",
                                part, ZipPackage.LineOf(element), tableName);
                        CheckStyleReference(element, TableNs + "style-name", part);
                        var rowNumber = totalRows + 1;
                        totalRows += CheckPositiveInteger(element, "number-rows-repeated", part) ?? 1;
                        ValidateRow(element, tableName, rowNumber, part);
                        break;

                    case "table-cell":
                    case "covered-table-cell":
                        if (element.Parent!.Name != TableNs + "table-row")
                            _issues.Error($"{Prefix}_CELL_MISPLACED",
                                $"<table:{element.Name.LocalName}> is inside <{element.Parent.Name.LocalName}>. Cells must be direct children of <table:table-row>.",
                                part, ZipPackage.LineOf(element), tableName);
                        break;
                }
            }

            if (totalRows > MaxRows)
                _issues.Warning($"{Prefix}_TOO_MANY_ROWS",
                    $"Sheet \"{tableName}\" declares {totalRows:N0} rows (including repeats), more than the {MaxRows:N0} rows spreadsheet applications support.", part, sheet: tableName);

            var hasData = table.Descendants(TableNs + "table-cell").Any(c =>
                c.Attribute(OfficeNs + "value-type") != null || c.Element(TextNs + "p") != null || c.Attribute(TableNs + "formula") != null);
            if (!hasData)
                _issues.Warning($"{Prefix}_SHEET_EMPTY", $"Sheet \"{tableName}\" contains no data.", part, ZipPackage.LineOf(table), tableName);
        }

        private void ValidateRow(XElement row, string tableName, long rowNumber, string part)
        {
            var totalColumns = 0L;
            var rowForIssues = rowNumber <= int.MaxValue ? (int)rowNumber : (int?)null;

            foreach (var cell in row.Elements())
            {
                if (cell.Name.Namespace != TableNs || cell.Name.LocalName is not ("table-cell" or "covered-table-cell"))
                {
                    if (cell.Name.Namespace == TableNs)
                        _issues.Error($"{Prefix}_UNEXPECTED_ELEMENT",
                            $"<table:table-row> may only contain table:table-cell and table:covered-table-cell, but contains <table:{cell.Name.LocalName}>.",
                            part, ZipPackage.LineOf(cell), tableName, row: rowForIssues);
                    continue;
                }

                var column = totalColumns + 1;
                var cellName = column <= MaxColumns && rowForIssues != null ? $"{XlsxValidator.ColumnName((int)column)}{rowNumber}" : null;
                totalColumns += CheckPositiveInteger(cell, "number-columns-repeated", part) ?? 1;
                CheckPositiveInteger(cell, "number-columns-spanned", part);
                CheckPositiveInteger(cell, "number-rows-spanned", part);
                CheckStyleReference(cell, TableNs + "style-name", part);
                ValidateCellValue(cell, part, tableName, cellName, rowForIssues);
            }

            if (totalColumns > MaxColumns)
                _issues.Warning($"{Prefix}_TOO_MANY_COLUMNS",
                    $"A row in sheet \"{tableName}\" declares {totalColumns:N0} columns (including repeats), more than the {MaxColumns:N0} columns spreadsheet applications support.",
                    part, ZipPackage.LineOf(row), tableName, row: rowForIssues);
        }

        private void ValidateCellValue(XElement cell, string part, string sheet, string? cellName, int? row)
        {
            var valueType = (string?)cell.Attribute(OfficeNs + "value-type");
            var location = cellName == null ? ZipPackage.LineOf(cell) : ZipPackage.At($"{sheet}.{cellName}", cell);

            var formula = (string?)cell.Attribute(TableNs + "formula");
            if (formula != null && formula.Contains("#REF!", StringComparison.Ordinal))
                _issues.Error($"{Prefix}_FORMULA_BROKEN_REFERENCE",
                    $"The formula \"{XlsxValidator.Truncate(formula)}\" contains #REF!, meaning it refers to a cell, range, or sheet that does not exist.",
                    part, location, sheet, cellName, row);
            if (!string.IsNullOrEmpty(formula))
                foreach (var missing in Services.FormulaTranslator.ReferencedSheets(Services.FormulaTranslator.FromOpenFormula(formula)).Where(n => !_tableNames.Contains(n)))
                    _issues.Error($"{Prefix}_FORMULA_UNKNOWN_SHEET",
                        $"The formula \"{XlsxValidator.Truncate(formula)}\" refers to sheet \"{missing}\", which does not exist in this document. Existing sheets: {string.Join(", ", _tableNames)}.",
                        part, location, sheet, cellName, row);

            if ((string?)cell.Attribute(CalcExtNs + "value-type") == "error")
            {
                var errorText = cell.Element(TextNs + "p")?.Value ?? "";
                _issues.Add(XlsxValidator.FormulaErrorSeverity(errorText), $"{Prefix}_CELL_FORMULA_ERROR",
                    $"The cell contains the error value {errorText}.{XlsxValidator.FormulaErrorHint(errorText)}", part, location, sheet, cellName, row);
            }

            if (valueType == null)
            {
                if (cell.Attribute(OfficeNs + "value") != null || cell.Attribute(OfficeNs + "date-value") != null || cell.Attribute(OfficeNs + "boolean-value") != null)
                    _issues.Warning($"{Prefix}_VALUE_WITHOUT_TYPE",
                        "The cell has a value attribute but no office:value-type, so applications will ignore the value. Add office:value-type (e.g. \"float\").",
                        part, location, sheet, cellName, row);
                return;
            }

            switch (valueType)
            {
                case "float":
                case "percentage":
                case "currency":
                    var number = (string?)cell.Attribute(OfficeNs + "value");
                    if (number == null)
                        _issues.Error($"{Prefix}_VALUE_MISSING",
                            $"The cell has office:value-type=\"{valueType}\" but no office:value attribute. Numeric cells store their number in office:value; the <text:p> is only the display text.",
                            part, location, sheet, cellName, row);
                    else if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                        _issues.Error($"{Prefix}_VALUE_NOT_NUMERIC",
                            $"office:value=\"{XlsxValidator.Truncate(number)}\" is not a number. Use a plain decimal such as 1234.5 (no thousands separators or currency symbols).",
                            part, location, sheet, cellName, row);
                    if (valueType == "currency" && cell.Attribute(OfficeNs + "currency") == null)
                        _issues.Warning($"{Prefix}_CURRENCY_MISSING", "The currency cell has no office:currency attribute (for example office:currency=\"USD\").", part, location, sheet, cellName, row);
                    break;

                case "date":
                    var date = (string?)cell.Attribute(OfficeNs + "date-value");
                    if (date == null)
                        _issues.Error($"{Prefix}_VALUE_MISSING", "The cell has office:value-type=\"date\" but no office:date-value attribute.", part, location, sheet, cellName, row);
                    else if (!IsXsdDate(date))
                        _issues.Error($"{Prefix}_DATE_INVALID",
                            $"office:date-value=\"{XlsxValidator.Truncate(date)}\" is not a valid date. Use 2024-01-31 or 2024-01-31T13:45:00.", part, location, sheet, cellName, row);
                    break;

                case "time":
                    var time = (string?)cell.Attribute(OfficeNs + "time-value");
                    if (time == null)
                        _issues.Error($"{Prefix}_VALUE_MISSING", "The cell has office:value-type=\"time\" but no office:time-value attribute.", part, location, sheet, cellName, row);
                    else if (!IsXsdDuration(time))
                        _issues.Error($"{Prefix}_TIME_INVALID",
                            $"office:time-value=\"{XlsxValidator.Truncate(time)}\" is not a valid duration. Use the ISO 8601 duration form, such as PT13H45M00S.", part, location, sheet, cellName, row);
                    break;

                case "boolean":
                    var boolean = (string?)cell.Attribute(OfficeNs + "boolean-value");
                    if (boolean == null)
                        _issues.Error($"{Prefix}_VALUE_MISSING", "The cell has office:value-type=\"boolean\" but no office:boolean-value attribute.", part, location, sheet, cellName, row);
                    else if (boolean is not ("true" or "false" or "1" or "0"))
                        _issues.Error($"{Prefix}_BOOLEAN_INVALID", $"office:boolean-value=\"{XlsxValidator.Truncate(boolean)}\" must be \"true\" or \"false\".", part, location, sheet, cellName, row);
                    break;

                case "string":
                    break;

                default:
                    _issues.Error($"{Prefix}_VALUE_TYPE_INVALID",
                        $"office:value-type=\"{valueType}\" is not valid. Use float, percentage, currency, date, time, boolean, or string.", part, location, sheet, cellName, row);
                    break;
            }
        }

        // ---------------------------------------------------------------- Helpers

        /// <summary>
        /// Checks that an optional table:* count attribute is a positive integer. Returns its value, or null if absent or invalid.
        /// </summary>
        private long? CheckPositiveInteger(XElement element, string localName, string part)
        {
            var attribute = element.Attribute(TableNs + localName);
            if (attribute == null)
                return null;

            if (long.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) && value > 0)
                return value;

            _issues.Warning($"{Prefix}_COUNT_INVALID",
                $"table:{localName}=\"{attribute.Value}\" must be a positive integer. Omit the attribute instead of leaving it empty.", part, ZipPackage.LineOf(attribute));
            return null;
        }

        private void CheckStyleReference(XElement element, XName attributeName, string part)
        {
            var attribute = element.Attribute(attributeName);
            if (attribute != null && !_styleNames.Contains(attribute.Value))
                _issues.Warning($"{Prefix}_STYLE_MISSING",
                    $"{attributeName.LocalName}=\"{attribute.Value}\" refers to a style that is not defined in content.xml or styles.xml. The default style will be used instead.",
                    part, ZipPackage.LineOf(attribute));
        }

        private static bool IsValidRangeAddress(string address)
        {
            // Form: [$]Sheet.[$]A[$]1:[$]Sheet.[$]B[$]2, where the sheet name may be quoted.
            foreach (var endpoint in address.Split(':'))
            {
                var dot = endpoint.LastIndexOf('.');
                if (dot < 0)
                    return false;
                var cell = endpoint[(dot + 1)..].Replace("$", "");
                if (XlsxValidator.ParseCellRef(cell) == null)
                    return false;
            }
            return true;
        }

        private static bool IsXsdDate(string value)
        {
            if (DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                return true;
            try
            {
                XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private static bool IsXsdDuration(string value)
        {
            try
            {
                XmlConvert.ToTimeSpan(value);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
            catch (OverflowException)
            {
                return true;
            }
        }
    }
}
