namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Settings that control how <see cref="SpreadsheetValidator"/> validates a file.
    /// </summary>
    public sealed class ValidationOptions
    {
        /// <summary>
        /// Maximum number of issues to collect before validation stops. Defaults to 100.
        /// </summary>
        public int MaxIssues { get; set; } = 100;

        /// <summary>
        /// Maximum combined uncompressed size, in bytes, of all entries in an XLSX or ODS package.
        /// Packages that declare a larger size are rejected without being decompressed, which protects against zip bombs.
        /// Defaults to 512 MB.
        /// </summary>
        public long MaxUncompressedBytes { get; set; } = 512L * 1024 * 1024;

        /// <summary>
        /// Delimiter expected in delimited files. When null (the default), it is inferred from the file extension
        /// (.tsv/.tab = tab, .psv = pipe) and otherwise defaults to <see cref="ColumnDelimeter.Comma"/>.
        /// </summary>
        public ColumnDelimeter? Delimiter { get; set; }

        /// <summary>
        /// Text encoding expected in delimited files. Defaults to <see cref="FileEncoding.UTF8"/>.
        /// </summary>
        public FileEncoding Encoding { get; set; } = FileEncoding.UTF8;

        /// <summary>
        /// When true, rows in a delimited file with a different number of fields than the first row are reported as errors.
        /// When false, they are reported as warnings. Defaults to true.
        /// </summary>
        public bool RequireConsistentColumnCount { get; set; } = true;
    }
}
