using OslSpreadsheet.Models;
using OslSpreadsheet.Services;

namespace OoxSpreadsheet
{
    public interface ISpreadsheet : IDisposable, IAsyncDisposable
    {
        oWorkbook Workbook { get; }

        Task<byte[]> GenerateCsvFileAsync();

        Task<byte[]> GenerateOdsFileAsync();

        Task<byte[]> GenerateXlsxFileAsync();

        Task<oWorkbook> ImportCsvFileAsync(byte[] file);

        Task<oWorkbook> ImportOdsFileAsync(byte[] file);

        Task<oWorkbook> ImportXlsxFileAsync(byte[] file);
    }

    public class Spreadsheet : ISpreadsheet
    {
        private oWorkbook _workbook;

        public Spreadsheet()
        {
            _workbook = new oWorkbook();
        }
        
        public oWorkbook Workbook  { get => _workbook; }


        public async Task<byte[]> GenerateCsvFileAsync()
        {
            IFileService _fileService = new DelimitedFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        public async Task<byte[]> GenerateOdsFileAsync()
        {
            IFileService _fileService = new OdsFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        public async Task<byte[]> GenerateXlsxFileAsync()
        {
            IFileService _fileService = new XlsxFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        public async Task<oWorkbook> ImportCsvFileAsync(byte[] file)
        {
            IFileService _fileService = new DelimitedFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

        public async Task<oWorkbook> ImportOdsFileAsync(byte[] file)
        {
            IFileService _fileService = new OdsFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

        public async Task<oWorkbook> ImportXlsxFileAsync(byte[] file)
        {
            IFileService _fileService = new XlsxFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

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