using System.IO.Compression;
using System.Text;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Services;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// A cell position is a row and column of at least 1. The model refuses anything else, and the importers
/// refuse references that don't name a real position instead of turning them into garbage coordinates.
/// </summary>
public class CellPositionTests
{
    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(-1, 1)]
    [InlineData(1, int.MinValue)]
    public void Constructor_RejectsNonPositivePositions(int row, int column)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new oCell(row, column));
    }

    [Fact]
    public void AddCell_RejectsNonPositivePositions()
    {
        var sheet = new oWorkbook().AddSheet("Data");

        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddCell(0, 1, "x"));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddCell(1, -5));
        Assert.Empty(sheet.Cells);
    }

    [Fact]
    public void Constructor_AcceptsGridCorners()
    {
        Assert.Equal((1, 1), (new oCell(1, 1).Row, new oCell(1, 1).Column));
        Assert.Equal((1_048_576, 16_384), (new oCell(1_048_576, 16_384).Row, new oCell(1_048_576, 16_384).Column));
    }

    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("Z1", 1, 26)]
    [InlineData("AA1", 1, 27)]
    [InlineData("XFD1048576", 1_048_576, 16_384)]
    [InlineData("B12", 12, 2)]
    public void TryParseCellRef_ValidReferences(string reference, int row, int col)
    {
        Assert.Equal((row, col), XlsxFileService.TryParseCellRef(reference));
    }

    [Theory]
    [InlineData("a1")]            // lowercase
    [InlineData("é1")]            // non-ASCII letter
    [InlineData("A0")]            // row 0
    [InlineData("A-1")]           // negative
    [InlineData("A")]             // no row
    [InlineData("1")]             // no column
    [InlineData("")]
    [InlineData("XFE1")]          // past the last column
    [InlineData("A1048577")]      // past the last row
    [InlineData("AAAA1")]         // four letters
    [InlineData("A12345678")]     // eight digits
    [InlineData("ABCDEFGHIJKLMNOP1")] // would overflow int
    [InlineData("A1.5")]
    [InlineData("A 1")]
    public void TryParseCellRef_InvalidReferences(string reference)
    {
        Assert.Null(XlsxFileService.TryParseCellRef(reference));
    }

    // ----- Imported files -----

    private static byte[] XlsxWithSheetXml(string sheetBody)
    {
        const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string rels = "http://schemas.openxmlformats.org/package/2006/relationships";
        var parts = new (string Path, string Content)[]
        {
            ("[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "</Types>"),
            ("_rels/.rels",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{rels}\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>"),
            ("xl/workbook.xml",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"{main}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"),
            ("xl/_rels/workbook.xml.rels",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{rels}\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                "</Relationships>"),
            ("xl/worksheets/sheet1.xml",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"{main}\">{sheetBody}</worksheet>")
        };

        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            foreach (var (path, content) in parts)
            {
                using var stream = archive.CreateEntry(path).Open();
                var bytes = Encoding.UTF8.GetBytes(content);
                stream.Write(bytes, 0, bytes.Length);
            }
        return ms.ToArray();
    }

    private static string Cell(string reference, string text) =>
        $"<c r=\"{reference}\" t=\"inlineStr\"><is><t>{text}</t></is></c>";

    [Theory]
    [InlineData("a1")]
    [InlineData("A0")]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("ABCDEFGHIJKLMNOP1")]
    public async Task XlsxImport_RejectsInvalidCellReference(string reference)
    {
        var bytes = XlsxWithSheetXml($"<sheetData><row r=\"1\">{Cell("A1", "ok")}{Cell(reference, "bad")}</row></sheetData>");

        using var spreadsheet = new Spreadsheet();
        var ex = await Assert.ThrowsAsync<SpreadsheetImportException>(() => spreadsheet.ImportXlsxFileAsync(bytes));

        Assert.Contains(reference, ex.Message);
    }

    [Fact]
    public async Task XlsxImport_IgnoresInvalidAutoFilterRange()
    {
        var bytes = XlsxWithSheetXml($"<sheetData><row r=\"1\">{Cell("A1", "ok")}</row></sheetData><autoFilter ref=\"a1:ZZZZ9\"/>");

        using var spreadsheet = new Spreadsheet();
        var sheet = (await spreadsheet.ImportXlsxFileAsync(bytes)).Sheets[0];

        Assert.Null(sheet.AutoFilterRange);
        Assert.Equal("ok", sheet.GetCell(1, 1)?.Value);
    }

    [Fact]
    public async Task XlsxImport_ReadsValidAutoFilterRange()
    {
        var bytes = XlsxWithSheetXml($"<sheetData><row r=\"1\">{Cell("A1", "ok")}</row></sheetData><autoFilter ref=\"A1:C9\"/>");

        using var spreadsheet = new Spreadsheet();
        var sheet = (await spreadsheet.ImportXlsxFileAsync(bytes)).Sheets[0];

        Assert.Equal((1, 1, 9, 3), sheet.AutoFilterRange);
    }

    /// <summary>
    /// Every cell the importers create satisfies the model's rule, so a file can never produce a cell the constructor would refuse.
    /// </summary>
    [Fact]
    public async Task OdsImport_QuotedSheetNameInFilterRange()
    {
        using var source = new Spreadsheet();
        var sheet = source.Workbook.AddSheet("My Data");
        sheet.AddCell(1, 1, "h");
        sheet.AddCell(2, 1, "v");
        sheet.SetAutoFilter();
        var bytes = await source.GenerateOdsFileAsync();

        using var spreadsheet = new Spreadsheet();
        var imported = (await spreadsheet.ImportOdsFileAsync(bytes)).Sheets[0];

        Assert.Equal((1, 1, 2, 1), imported.AutoFilterRange);
    }
}
