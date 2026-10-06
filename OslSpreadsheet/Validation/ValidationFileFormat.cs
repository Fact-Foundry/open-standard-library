namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// File formats recognized by <see cref="SpreadsheetValidator"/>.
    /// </summary>
    public enum ValidationFileFormat
    {
        /// <summary>
        /// The format could not be determined.
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// Office Open XML spreadsheet (.xlsx, .xlsm).
        /// </summary>
        Xlsx,

        /// <summary>
        /// OpenDocument spreadsheet (.ods).
        /// </summary>
        Ods,

        /// <summary>
        /// Comma, tab, pipe, or ASCII delimited text (.csv, .tsv, .psv, .txt).
        /// </summary>
        Delimited
    }
}
