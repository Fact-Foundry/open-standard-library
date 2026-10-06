namespace OslSpreadsheet.Models
{
    /// <summary>
    /// Represents a file in memory for ZIP archive construction.
    /// </summary>
    internal class InMemoryFile
    {
        /// <summary>
        /// The path and name of the file within the archive.
        /// </summary>
        public required string FileName { get; set; }

        /// <summary>
        /// The file content as a byte array.
        /// </summary>
        public required byte[] Content { get; set; }

        /// <summary>
        /// When true, the file is stored uncompressed in the archive.
        /// </summary>
        public bool Store { get; set; }
    }
}
