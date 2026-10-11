using System.IO.Compression;
using System.Text;
using OoxSpreadsheet;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests that every way a malformed file can fail import surfaces as SpreadsheetImportException.
/// </summary>
public class ImportExceptionTests
{
    private static byte[] Zip(params (string Path, string Content)[] entries)
    {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            foreach (var (path, content) in entries)
            {
                using var stream = archive.CreateEntry(path).Open();
                var bytes = Encoding.UTF8.GetBytes(content);
                stream.Write(bytes, 0, bytes.Length);
            }
        return ms.ToArray();
    }

    /// <summary>
    /// Bytes that aren't a ZIP archive at all.
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task NotAZip_ThrowsImportException(bool ods)
    {
        var bytes = "Name,Score\nAlice,95\n"u8.ToArray();

        using var spreadsheet = new Spreadsheet();
        var ex = await Assert.ThrowsAsync<SpreadsheetImportException>(() => ods ? spreadsheet.ImportOdsFileAsync(bytes) : spreadsheet.ImportXlsxFileAsync(bytes));

        Assert.IsType<InvalidDataException>(ex.InnerException);
        Assert.Contains(ods ? "ODS" : "XLSX", ex.Message);
    }

    /// <summary>
    /// A ZIP archive without the part that identifies the format.
    /// </summary>
    [Fact]
    public async Task MissingRequiredPart_ThrowsImportException()
    {
        var notXlsx = Zip(("hello.txt", "hi"));

        using var spreadsheet = new Spreadsheet();
        var ex = await Assert.ThrowsAsync<SpreadsheetImportException>(() => spreadsheet.ImportXlsxFileAsync(notXlsx));
        Assert.Contains("xl/workbook.xml", ex.Message);

        using var odsImporter = new Spreadsheet();
        var odsEx = await Assert.ThrowsAsync<SpreadsheetImportException>(() => odsImporter.ImportOdsFileAsync(notXlsx));
        Assert.Contains("content.xml", odsEx.Message);
    }

    /// <summary>
    /// Malformed XML inside a part surfaces with the XmlException as the inner exception.
    /// </summary>
    [Fact]
    public async Task MalformedXml_ThrowsImportExceptionWithInner()
    {
        var ods = Zip(
            ("mimetype", "application/vnd.oasis.opendocument.spreadsheet"),
            ("content.xml", "<office:document-content><unclosed>"));

        using var spreadsheet = new Spreadsheet();
        var ex = await Assert.ThrowsAsync<SpreadsheetImportException>(() => spreadsheet.ImportOdsFileAsync(ods));

        Assert.IsType<System.Xml.XmlException>(ex.InnerException);
        Assert.Contains("SpreadsheetValidator", ex.Message);
    }

    /// <summary>
    /// A relationship element missing its Id or Target no longer crashes the import; the sheet it would have pointed to is skipped.
    /// </summary>
    [Fact]
    public async Task Xlsx_RelationshipWithoutId_IsSkipped()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "x");
        var original = await spreadsheet.GenerateXlsxFileAsync();

        using var input = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read);
        using var ms = new MemoryStream();
        using (var output = new ZipArchive(ms, ZipArchiveMode.Create, true))
            foreach (var entry in input.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                var text = reader.ReadToEnd();
                if (entry.FullName == "xl/_rels/workbook.xml.rels")
                    text = text.Replace("Id=\"rId1\"", "");
                using var stream = output.CreateEntry(entry.FullName).Open();
                var bytes = Encoding.UTF8.GetBytes(text);
                stream.Write(bytes, 0, bytes.Length);
            }

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(ms.ToArray());

        Assert.Empty(workbook.Sheets);
    }

    /// <summary>
    /// Limit violations are SpreadsheetImportExceptions too, so one catch covers both.
    /// </summary>
    [Fact]
    public async Task LimitExceeded_IsAnImportException()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        for (int c = 1; c <= 10; c++) sheet.AddCell(1, c, "v");
        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var ex = await Assert.ThrowsAsync<ImportLimitExceededException>(() => importer.ImportXlsxFileAsync(bytes, new ImportOptions { MaxCells = 5 }));

        Assert.IsAssignableFrom<SpreadsheetImportException>(ex);
    }

    /// <summary>
    /// A null file argument is a caller bug, not a malformed file, and is reported as such.
    /// </summary>
    [Fact]
    public async Task NullFile_ThrowsArgumentNullException()
    {
        using var spreadsheet = new Spreadsheet();
        await Assert.ThrowsAsync<ArgumentNullException>(() => spreadsheet.ImportXlsxFileAsync(null!));
    }
}
