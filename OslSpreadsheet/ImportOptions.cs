namespace OoxSpreadsheet
{
    /// <summary>
    /// Limits applied when importing a file, so a small malicious or malformed file can't exhaust memory.
    /// Exceeding a limit throws <see cref="ImportLimitExceededException"/>.
    /// </summary>
    public sealed class ImportOptions
    {
        /// <summary>
        /// Maximum combined uncompressed size, in bytes, of all entries in an XLSX or ODS package.
        /// Parsed XML costs roughly ten times its byte size in memory, so keep this well below available memory.
        /// Defaults to 64 MB.
        /// </summary>
        public long MaxUncompressedBytes { get; set; } = 64L * 1024 * 1024;

        /// <summary>
        /// Maximum number of cells imported across all sheets. ODS files can declare a cell repeated millions of times
        /// in a few bytes, so this bounds the in-memory model rather than the file. Defaults to 10,000,000.
        /// </summary>
        public int MaxCells { get; set; } = 10_000_000;
    }

    /// <summary>
    /// Thrown by the Import methods when a file can't be read: it is not the expected format, is corrupt or truncated,
    /// contains malformed XML, or exceeds an import limit (see <see cref="ImportLimitExceededException"/>).
    /// <see cref="Exception.InnerException"/> holds the underlying parser error when there is one.
    /// Catching this one type covers every malformed-input failure; use <c>SpreadsheetValidator</c> first for a detailed report.
    /// </summary>
    public class SpreadsheetImportException : Exception
    {
        internal SpreadsheetImportException(string message, Exception? innerException = null) : base(message, innerException) { }
    }

    /// <summary>
    /// Thrown when a file exceeds one of the <see cref="ImportOptions"/> limits.
    /// </summary>
    public sealed class ImportLimitExceededException : SpreadsheetImportException
    {
        internal ImportLimitExceededException(string message) : base(message) { }
    }
}
