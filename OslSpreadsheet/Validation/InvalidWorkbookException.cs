namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Thrown by the Generate methods when the workbook would produce an invalid file.
    /// <see cref="Result"/> lists the problems; call <c>oWorkbook.Validate()</c> beforehand to check without throwing.
    /// </summary>
    public sealed class InvalidWorkbookException : Exception
    {
        internal InvalidWorkbookException(ValidationResult result)
            : base(result.ToString())
        {
            Result = result;
        }

        /// <summary>
        /// The validation result describing each problem, with sheet, cell, and row where applicable.
        /// </summary>
        public ValidationResult Result { get; }
    }
}
