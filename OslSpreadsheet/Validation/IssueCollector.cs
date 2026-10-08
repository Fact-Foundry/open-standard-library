namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Thrown internally to stop validation once the issue limit is reached.
    /// </summary>
    internal sealed class IssueLimitReachedException : Exception { }

    /// <summary>
    /// Accumulates issues during validation and enforces <see cref="ValidationOptions.MaxIssues"/>.
    /// </summary>
    internal sealed class IssueCollector
    {
        private readonly List<ValidationIssue> _issues = new();
        private readonly int _maxIssues;

        internal IssueCollector(ValidationOptions options)
        {
            _maxIssues = Math.Max(1, options.MaxIssues);
        }

        internal bool Truncated { get; private set; }

        internal bool HasErrors => _issues.Any(i => i.Severity == ValidationSeverity.Error);

        internal void Error(string code, string message, string? part = null, string? location = null,
            string? sheet = null, string? cell = null, int? row = null) =>
            Add(ValidationSeverity.Error, code, message, part, location, sheet, cell, row);

        internal void Warning(string code, string message, string? part = null, string? location = null,
            string? sheet = null, string? cell = null, int? row = null) =>
            Add(ValidationSeverity.Warning, code, message, part, location, sheet, cell, row);

        internal void Add(ValidationSeverity severity, string code, string message, string? part = null, string? location = null,
            string? sheet = null, string? cell = null, int? row = null)
        {
            if (_issues.Count >= _maxIssues)
            {
                Truncated = true;
                throw new IssueLimitReachedException();
            }

            _issues.Add(new ValidationIssue(severity, code, message, part, location, sheet, cell, row));
        }

        internal ValidationResult ToResult(ValidationFileFormat format, string subject = "file") => new(format, _issues.ToList(), Truncated, subject);
    }
}
