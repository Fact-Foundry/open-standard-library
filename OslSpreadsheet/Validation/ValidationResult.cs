using System.Text;

namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// The outcome of validating a file with <see cref="SpreadsheetValidator"/>.
    /// </summary>
    public sealed class ValidationResult
    {
        private readonly string _subject;

        internal ValidationResult(ValidationFileFormat format, IReadOnlyList<ValidationIssue> issues, bool truncated, string subject = "file")
        {
            Format = format;
            Issues = issues;
            Truncated = truncated;
            _subject = subject;
        }

        /// <summary>
        /// The format the file was validated as.
        /// </summary>
        public ValidationFileFormat Format { get; }

        /// <summary>
        /// True when no errors were found. Warnings do not make a file invalid.
        /// </summary>
        public bool IsValid => !Issues.Any(i => i.Severity == ValidationSeverity.Error);

        /// <summary>
        /// All issues found, in the order they were encountered.
        /// </summary>
        public IReadOnlyList<ValidationIssue> Issues { get; }

        /// <summary>
        /// Issues with <see cref="ValidationSeverity.Error"/> severity.
        /// </summary>
        public IEnumerable<ValidationIssue> Errors => Issues.Where(i => i.Severity == ValidationSeverity.Error);

        /// <summary>
        /// Issues with <see cref="ValidationSeverity.Warning"/> severity.
        /// </summary>
        public IEnumerable<ValidationIssue> Warnings => Issues.Where(i => i.Severity == ValidationSeverity.Warning);

        /// <summary>
        /// True when validation stopped early because <see cref="ValidationOptions.MaxIssues"/> was reached.
        /// </summary>
        public bool Truncated { get; }

        /// <summary>
        /// Plain-text report of the result, suitable for showing to a user or returning to an LLM so it can correct the file.
        /// </summary>
        public override string ToString()
        {
            var name = FormatName(Format);
            var errorCount = Errors.Count();
            var warningCount = Warnings.Count();

            var sb = new StringBuilder();
            if (_subject == "workbook")
                sb.Append(IsValid ? $"The workbook is valid for {name} output" : $"The workbook is NOT valid for {name} output");
            else if (IsValid)
                sb.Append($"The file is a valid {name} file");
            else
                sb.Append($"The file is NOT a valid {name} file");

            if (Issues.Count == 0)
                return sb.Append('.').ToString();

            sb.Append($" ({Plural(errorCount, "error")}, {Plural(warningCount, "warning")}):");
            foreach (var issue in Issues)
                sb.AppendLine().Append("- ").Append(issue);

            if (Truncated)
                sb.AppendLine().Append("- Validation stopped after the maximum number of issues; more may exist.");

            return sb.ToString();
        }

        private static string FormatName(ValidationFileFormat format) => format switch
        {
            ValidationFileFormat.Xlsx => "XLSX",
            ValidationFileFormat.Ods => "ODS",
            ValidationFileFormat.Delimited => "delimited text",
            _ => "spreadsheet"
        };

        private static string Plural(int count, string word) => count == 1 ? $"1 {word}" : $"{count} {word}s";
    }
}
