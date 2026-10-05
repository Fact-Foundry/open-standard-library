namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// A single problem found while validating a file.
    /// </summary>
    public sealed class ValidationIssue
    {
        internal ValidationIssue(ValidationSeverity severity, string code, string message, string? part, string? location,
            string? sheet, string? cell, int? row)
        {
            Severity = severity;
            Code = code;
            Message = message;
            Part = part;
            Location = location;
            Sheet = sheet;
            Cell = cell;
            Row = row;
        }

        /// <summary>
        /// Whether the issue makes the file invalid (<see cref="ValidationSeverity.Error"/>) or is advisory.
        /// </summary>
        public ValidationSeverity Severity { get; }

        /// <summary>
        /// Stable machine-readable identifier for the kind of issue, e.g. <c>XLSX_CELL_NOT_NUMERIC</c>.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Human-readable description of the problem and, where possible, how to fix it.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// The package part (zip entry) the issue was found in, e.g. <c>xl/worksheets/sheet1.xml</c>. Null for delimited files.
        /// </summary>
        public string? Part { get; }

        /// <summary>
        /// Where in the part the issue was found, e.g. <c>cell B3</c>, <c>line 12, position 4</c>, or <c>row 5</c>.
        /// </summary>
        public string? Location { get; }

        /// <summary>
        /// Name of the sheet the issue was found in, when it applies to a specific sheet. Null for delimited files.
        /// </summary>
        public string? Sheet { get; }

        /// <summary>
        /// A1-style address of the cell the issue was found in, e.g. <c>B3</c>, when it applies to a specific cell.
        /// </summary>
        public string? Cell { get; }

        /// <summary>
        /// 1-based row number the issue was found in, when it applies to a specific row.
        /// For delimited files this is the record number, which can differ from the line number when values contain line breaks.
        /// </summary>
        public int? Row { get; }

        /// <summary>
        /// Formats the issue as a single line, e.g. <c>[Error] xl/worksheets/sheet1.xml, cell B3: ...</c>.
        /// </summary>
        public override string ToString()
        {
            var where = string.Join(", ", new[] { Part, Location }.Where(s => !string.IsNullOrEmpty(s)));
            return where.Length == 0
                ? $"[{Severity}] {Code}: {Message}"
                : $"[{Severity}] {Code} ({where}): {Message}";
        }
    }
}
