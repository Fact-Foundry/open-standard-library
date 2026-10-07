using OslSpreadsheet.Models;
using System.Text;

namespace OslSpreadsheet.Services
{
    internal class DelimitedFileService : IFileService
    {
        private readonly ColumnDelimeter _importDelimiter;
        private readonly FileEncoding _importEncoding;

        /// <summary>
        /// Creates the service. The delimiter and encoding apply to import; generation uses the workbook's settings.
        /// </summary>
        public DelimitedFileService(ColumnDelimeter importDelimiter = ColumnDelimeter.Comma, FileEncoding importEncoding = FileEncoding.UTF8)
        {
            _importDelimiter = importDelimiter;
            _importEncoding = importEncoding;
        }

        /// <summary>
        /// Converts an oWorkbook object to a delimited file byte array.
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public Task<byte[]> GenerateFileAsync(oWorkbook workbook)
        {
            if (workbook.Sheets.Count == 0) throw new ArgumentException("Workbook requires at least 1 sheet to convert to a delimited file.");

            var delimiter = workbook.ColumnDelimeter;
            var colDelimiter = DelimitedParser.GetDelimiter(delimiter);
            var nlDelimiter = GetRowDelimiter(delimiter);

            var cells = workbook.Sheets[0].Cells;
            var cols = cells.Max(x => x.Column);
            var rows = cells.Max(x => x.Row);
            var lookup = cells.GroupBy(c => (c.Row, c.Column)).ToDictionary(g => g.Key, g => g.Last().Value);

            var result = new StringBuilder();

            // Have to start at 1 instead of 0
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    if (c > 1)
                        result.Append(colDelimiter);
                    var value = lookup.TryGetValue((r, c), out var v) ? v ?? "" : "";
                    result.Append(DelimitedParser.FormatValue(value, delimiter));
                }
                result.Append(nlDelimiter);
            }

            return Task.FromResult(GetEncoding(workbook.FileEncoding).GetBytes(result.ToString()));
        }

        private static string GetRowDelimiter(ColumnDelimeter delimeter)
        {
            switch (delimeter)
            {
                case ColumnDelimeter.ASCII:
                    return "\u001E";
                default:
                case ColumnDelimeter.Comma:
                case ColumnDelimeter.Pipe:
                case ColumnDelimeter.Tab:
                    return Environment.NewLine;
            }
        }

        /// <summary>
        /// Converts a delimited file byte array to an oWorkbook object, using the delimiter and encoding given to the constructor.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public async Task<oWorkbook> GenerateModel(byte[] file)
        {
            oWorkbook workbook = new()
            {
                ColumnDelimeter = _importDelimiter,
                FileEncoding = _importEncoding
            };

            var sheet1 = await workbook.AddSheetAsync();

            if (file == null || file.Length == 0)
                return workbook;

            using var reader = new StreamReader(new MemoryStream(file), GetEncoding(_importEncoding));

            int r = 0;
            await foreach (var values in DelimitedParser.ParseAsync(reader, _importDelimiter))
            {
                r++;
                for (int c = 0; c < values.Length; c++)
                    sheet1.AddCell(r, c + 1, values[c]);
            }

            return workbook;
        }

        /// <summary>
        /// Maps a FileEncoding enum value to its corresponding System.Text.Encoding instance.
        /// </summary>
        /// <param name="fileEncoding"></param>
        /// <returns></returns>
        private static Encoding GetEncoding(FileEncoding fileEncoding) => fileEncoding switch
        {
            FileEncoding.ASCII => Encoding.ASCII,
            FileEncoding.Unicode => Encoding.Unicode,
            FileEncoding.UTF32 => Encoding.UTF32,
            _ => Encoding.UTF8
        };

        public void Dispose() { }

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
