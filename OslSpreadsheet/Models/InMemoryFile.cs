namespace OslSpreadsheet.Models
{
    internal class InMemoryFile
    {
        public string FileName { get; set; }

        public byte[] Content { get; set; }

        public bool Store { get; set; }
    }
}
