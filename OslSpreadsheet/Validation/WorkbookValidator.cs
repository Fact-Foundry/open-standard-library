using System.Globalization;
using OslSpreadsheet.Models;
using OslSpreadsheet.Services;

namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Checks an in-memory <see cref="oWorkbook"/> for problems that would produce an invalid or lossy file
    /// before any file is generated. Rules are format-aware, since XLSX, ODS, and delimited files have different limits.
    /// </summary>
    internal static class WorkbookValidator
    {
        private const string Prefix = "WB";
        private const int MaxRows = 1_048_576;
        private const int MaxColumns = 16_384;
        private const int MaxTextLength = 32_767;

        internal static ValidationResult Validate(oWorkbook workbook, ValidationFileFormat format, ValidationOptions? options = null)
        {
            options ??= new ValidationOptions();
            var issues = new IssueCollector(options);

            try
            {
                Run(workbook, format, issues);
            }
            catch (IssueLimitReachedException)
            {
                // The collector already recorded that the result is truncated.
            }

            return issues.ToResult(format, subject: "workbook");
        }

        private static void Run(oWorkbook workbook, ValidationFileFormat format, IssueCollector issues)
        {
            if (workbook.Sheets.Count == 0)
            {
                issues.Error($"{Prefix}_NO_SHEETS", "The workbook has no sheets. Add at least one sheet before generating a file.");
                return;
            }

            if (format == ValidationFileFormat.Delimited && workbook.Sheets.Count > 1)
                issues.Warning($"{Prefix}_DELIMITED_MULTIPLE_SHEETS",
                    $"The workbook has {workbook.Sheets.Count} sheets, but a delimited file can only hold one. Only \"{workbook.Sheets[0].SheetName}\" will be written.");

            var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sheet in workbook.Sheets)
                ValidateSheetName(sheet, format, sheetNames, issues);

            foreach (var sheet in workbook.Sheets)
                ValidateSheet(sheet, format, sheetNames, workbook, issues);
        }

        private static void ValidateSheetName(oSpreadsheet sheet, ValidationFileFormat format, HashSet<string> seen, IssueCollector issues)
        {
            var name = sheet.SheetName;

            if (string.IsNullOrWhiteSpace(name))
            {
                issues.Error($"{Prefix}_SHEET_NAME_INVALID", "A sheet has a blank name.", sheet: name);
                return;
            }

            // Delimited files don't store sheet names, so only XLSX and ODS care about their form.
            var problem = format switch
            {
                ValidationFileFormat.Xlsx => XlsxValidator.ValidateSheetName(name),
                ValidationFileFormat.Ods => OdsValidator.ValidateTableName(name),
                _ => null
            };
            if (problem != null)
                issues.Error($"{Prefix}_SHEET_NAME_INVALID", $"Sheet name \"{name}\" {problem}", sheet: name);

            if (!seen.Add(name) && format != ValidationFileFormat.Delimited)
                issues.Error($"{Prefix}_SHEET_NAME_DUPLICATE", $"Sheet name \"{name}\" is used more than once. Sheet names must be unique (case-insensitive).", sheet: name);
        }

        private static void ValidateSheet(oSpreadsheet sheet, ValidationFileFormat format, HashSet<string> sheetNames, oWorkbook workbook, IssueCollector issues)
        {
            var name = sheet.SheetName;

            if (sheet.FreezeRows < 0 || sheet.FreezeColumns < 0)
                issues.Error($"{Prefix}_FREEZE_INVALID", $"FreezeRows ({sheet.FreezeRows}) and FreezeColumns ({sheet.FreezeColumns}) can't be negative.", sheet: name);

            if (sheet.AutoFilterRange is var (sr, sc, er, ec)
                && (sr < 1 || sc < 1 || er < sr || ec < sc || er > MaxRows || ec > MaxColumns))
                issues.Error($"{Prefix}_AUTOFILTER_INVALID",
                    $"The auto-filter range (rows {sr}-{er}, columns {sc}-{ec}) is invalid. Start must be at least 1 and end must not be before start.", sheet: name);

            foreach (var cell in sheet.Cells)
                ValidateCell(cell, name, format, sheetNames, issues);
        }

        private static void ValidateCell(oCell cell, string sheetName, ValidationFileFormat format, HashSet<string> sheetNames, IssueCollector issues)
        {
            var cellName = cell.Column >= 1 && cell.Column <= MaxColumns ? $"{XlsxValidator.ColumnName(cell.Column)}{cell.Row}" : $"R{cell.Row}C{cell.Column}";
            var location = $"{sheetName}!{cellName}";
            var value = cell.Value ?? "";

            if (cell.Row < 1 || cell.Column < 1 || cell.Row > MaxRows || cell.Column > MaxColumns)
            {
                issues.Error($"{Prefix}_CELL_POSITION_INVALID",
                    $"Cell position row {cell.Row}, column {cell.Column} is out of range. Rows run from 1 to {MaxRows:N0} and columns from 1 to {MaxColumns:N0}.",
                    location: location, sheet: sheetName, cell: cellName, row: cell.Row);
                return;
            }

            // A formula cell with no cached result legitimately has an empty value whatever its type
            var hasFormula = FormulaTranslator.Normalize(cell.Formula) != null;

            switch (hasFormula && value.Length == 0 ? CellValueType.String : cell.ValueType)
            {
                case CellValueType.Float:
                    if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                        issues.Error($"{Prefix}_CELL_NOT_NUMERIC",
                            $"The cell is typed Float but its value \"{XlsxValidator.Truncate(value)}\" is not a number. Use a plain decimal such as 1234.5 (no thousands separators or currency symbols), or change ValueType to String.",
                            location: location, sheet: sheetName, cell: cellName, row: cell.Row);
                    break;

                case CellValueType.Int64:
                    if (!long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                        issues.Error($"{Prefix}_CELL_NOT_NUMERIC",
                            $"The cell is typed Int64 but its value \"{XlsxValidator.Truncate(value)}\" is not a whole number.",
                            location: location, sheet: sheetName, cell: cellName, row: cell.Row);
                    break;

                case CellValueType.Boolean:
                    if (value is not ("true" or "false" or "True" or "False" or "TRUE" or "FALSE" or "1" or "0"))
                        issues.Error($"{Prefix}_CELL_NOT_BOOLEAN",
                            $"The cell is typed Boolean but its value \"{XlsxValidator.Truncate(value)}\" is not true or false.",
                            location: location, sheet: sheetName, cell: cellName, row: cell.Row);
                    break;

                case CellValueType.DateTime:
                    if (!DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                        issues.Error($"{Prefix}_CELL_NOT_DATETIME",
                            $"The cell is typed DateTime but its value \"{XlsxValidator.Truncate(value)}\" is not a recognized date. Use ISO 8601, such as 2024-01-31 or 2024-01-31T13:45:00.",
                            location: location, sheet: sheetName, cell: cellName, row: cell.Row);
                    break;

                default:
                    if (format != ValidationFileFormat.Delimited && value.Length > MaxTextLength)
                        issues.Add(format == ValidationFileFormat.Xlsx ? ValidationSeverity.Error : ValidationSeverity.Warning,
                            $"{Prefix}_TEXT_TOO_LONG",
                            $"The cell text is {value.Length:N0} characters long; spreadsheet cells are limited to {MaxTextLength:N0} characters.",
                            location: location, sheet: sheetName, cell: cellName, row: cell.Row);
                    break;
            }

            if (cell.Style?.NumberFormat is { Length: > 0 } numberFormat)
            {
                if (FindControlChar(numberFormat) is char badFormat)
                    issues.Error($"{Prefix}_NUMBER_FORMAT_INVALID",
                        $"The number format contains the control character U+{(int)badFormat:X4}.",
                        location: location, sheet: sheetName, cell: cellName, row: cell.Row);
                else if (format == ValidationFileFormat.Ods && NumberFormatTranslator.ToOdsDataStyle(numberFormat, "x") == null)
                    issues.Warning($"{Prefix}_NUMBER_FORMAT_UNSUPPORTED",
                        $"The number format \"{XlsxValidator.Truncate(numberFormat)}\" can't be translated to an ODS data style, so the cell will use the default format. Supported: decimals, thousands separators, percentages, currency symbols, text literals, and date/time patterns.",
                        location: location, sheet: sheetName, cell: cellName, row: cell.Row);
            }

            if (FindControlChar(value) is char bad)
                issues.Error($"{Prefix}_TEXT_CONTROL_CHAR",
                    $"The cell value contains the control character U+{(int)bad:X4}, which can't be stored in {FormatName(format)}. Remove it or replace it with a space.",
                    location: location, sheet: sheetName, cell: cellName, row: cell.Row);

            if (format == ValidationFileFormat.Delimited)
                return;

            var formula = FormulaTranslator.Normalize(cell.Formula);
            if (formula == null)
                return;

            if (FindControlChar(formula) is char badFormula)
                issues.Error($"{Prefix}_TEXT_CONTROL_CHAR",
                    $"The formula contains the control character U+{(int)badFormula:X4}, which can't be stored in {FormatName(format)}.",
                    location: location, sheet: sheetName, cell: cellName, row: cell.Row);

            foreach (var missing in FormulaTranslator.ReferencedSheets(formula).Where(s => !sheetNames.Contains(s)))
                issues.Error($"{Prefix}_FORMULA_UNKNOWN_SHEET",
                    $"The formula \"{XlsxValidator.Truncate(formula)}\" refers to sheet \"{missing}\", which is not in the workbook. Existing sheets: {string.Join(", ", sheetNames)}.",
                    location: location, sheet: sheetName, cell: cellName, row: cell.Row);
        }

        /// <summary>
        /// Returns the first character that XML 1.0 can't represent (C0 controls other than tab, LF, CR), or null.
        /// For delimited files, the ASCII unit and record separators are also disallowed since they delimit the file.
        /// </summary>
        private static char? FindControlChar(string text)
        {
            foreach (var c in text)
                if (c < 0x20 && c is not ('\t' or '\n' or '\r'))
                    return c;
            return null;
        }

        private static string FormatName(ValidationFileFormat format) => format switch
        {
            ValidationFileFormat.Xlsx => "XLSX",
            ValidationFileFormat.Ods => "ODS",
            ValidationFileFormat.Delimited => "a delimited file",
            _ => "the file"
        };
    }
}
