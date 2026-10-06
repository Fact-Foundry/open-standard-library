using OslSpreadsheet.Models;
using OslSpreadsheet.Services;
using System.Text;

namespace OoxSpreadsheet
{
    /// <summary>
    /// Interface for generating and importing spreadsheet files in multiple formats.
    /// </summary>
    public interface ISpreadsheet : IDisposable, IAsyncDisposable
    {
        /// <summary>
        /// The workbook containing all sheets and data.
        /// </summary>
        oWorkbook Workbook { get; }

        /// <summary>
        /// Headers from the most recent ReadCsvRowsAsync call when hasHeaderRow was true.
        /// </summary>
        string[]? CsvHeaders { get; }

        /// <summary>
        /// Generates a CSV file from the workbook.
        /// </summary>
        Task<byte[]> GenerateCsvFileAsync();

        /// <summary>
        /// Generates an ODS (OpenDocument Spreadsheet) file from the workbook.
        /// </summary>
        Task<byte[]> GenerateOdsFileAsync();

        /// <summary>
        /// Generates an XLSX (Office Open XML) file from the workbook.
        /// </summary>
        Task<byte[]> GenerateXlsxFileAsync();

        /// <summary>
        /// Imports a CSV file into the workbook.
        /// </summary>
        Task<oWorkbook> ImportCsvFileAsync(byte[] file);

        /// <summary>
        /// Imports an ODS file into the workbook.
        /// </summary>
        Task<oWorkbook> ImportOdsFileAsync(byte[] file);

        /// <summary>
        /// Imports an XLSX file into the workbook.
        /// </summary>
        Task<oWorkbook> ImportXlsxFileAsync(byte[] file);

        /// <summary>
        /// Reads CSV rows one at a time from a stream without loading the entire file into memory.
        /// When hasHeaderRow is true, the first row is consumed as headers (available via CsvHeaders) and not yielded.
        /// </summary>
        IAsyncEnumerable<string[]> ReadCsvRowsAsync(Stream stream, bool hasHeaderRow = false, int? rowLimit = null);
    }

    /// <summary>
    /// Spreadsheet implementation for generating and importing ODS, XLSX, and delimited files.
    /// </summary>
    public class Spreadsheet : ISpreadsheet
    {
        private oWorkbook _workbook;

        public Spreadsheet()
        {
            _workbook = new oWorkbook();
        }

        /// <inheritdoc />
        public oWorkbook Workbook  { get => _workbook; }

        /// <inheritdoc />
        public string[]? CsvHeaders { get; private set; }

        /// <inheritdoc />
        public async Task<byte[]> GenerateCsvFileAsync()
        {
            IFileService _fileService = new DelimitedFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        /// <inheritdoc />
        public async Task<byte[]> GenerateOdsFileAsync()
        {
            IFileService _fileService = new OdsFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        /// <inheritdoc />
        public async Task<byte[]> GenerateXlsxFileAsync()
        {
            IFileService _fileService = new XlsxFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        /// <inheritdoc />
        public async Task<oWorkbook> ImportCsvFileAsync(byte[] file)
        {
            IFileService _fileService = new DelimitedFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

        /// <inheritdoc />
        public async Task<oWorkbook> ImportOdsFileAsync(byte[] file)
        {
            IFileService _fileService = new OdsFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

        /// <inheritdoc />
        public async Task<oWorkbook> ImportXlsxFileAsync(byte[] file)
        {
            IFileService _fileService = new XlsxFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

        /// <inheritdoc />
        public async IAsyncEnumerable<string[]> ReadCsvRowsAsync(Stream stream, bool hasHeaderRow = false, int? rowLimit = null)
        {
            CsvHeaders = null;
            var encoding = GetEncoding(_workbook.FileEncoding);
            using var reader = new StreamReader(stream, encoding);

            int rowCount = 0;
            bool isFirstRow = true;

            string? line;
            while ((line = await reader.ReadLineAsync()) != null)
            {
                if (string.IsNullOrEmpty(line)) continue;

                var values = ParseCsvLine(line);

                if (isFirstRow && hasHeaderRow)
                {
                    CsvHeaders = values;
                    isFirstRow = false;
                    continue;
                }

                isFirstRow = false;

                if (rowLimit.HasValue && rowCount >= rowLimit.Value)
                    yield break;

                rowCount++;
                yield return values;
            }
        }

        private static string[] ParseCsvLine(string line)
        {
            var normalized = line.Replace("\", \"", "\",\"").Replace("\" ,\"", "\",\"");
            var cols = normalized.Split("\",\"");

            if (cols[0].StartsWith("\""))
                cols[0] = cols[0][1..];

            var last = cols.Length - 1;
            if (cols[last].EndsWith("\""))
                cols[last] = cols[last][..^1];

            for (int i = 0; i < cols.Length; i++)
                cols[i] = cols[i].Replace("\"\"", "\"");

            return cols;
        }

        private static Encoding GetEncoding(FileEncoding fileEncoding) => fileEncoding switch
        {
            FileEncoding.ASCII => Encoding.ASCII,
            FileEncoding.Unicode => Encoding.Unicode,
            FileEncoding.UTF32 => Encoding.UTF32,
            _ => Encoding.UTF8
        };

        public async ValueTask DisposeAsync()
        {
            // Todo: Add implementation
            await DisposeAsyncCore();

            Dispose(disposing: false);
#pragma warning disable CA1816 // Dispose methods should call SuppressFinalize
            GC.SuppressFinalize(this);
#pragma warning restore CA1816 // Dispose methods should call SuppressFinalize
        }

        protected virtual async ValueTask DisposeAsyncCore()
        {
            await Task.Delay(1);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // dispose of anything necessary here
            }
        }
    }
}