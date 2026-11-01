using System.Diagnostics;
using System.IO.Compression;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Generation must cost in proportion to the cells that exist, not the grid area between them: a sheet with
/// A1 and XFD1048576 has two cells, not 17 billion positions.
/// </summary>
public class SparseSheetTests
{
    private const int MaxColumns = 16384;
    private const int MaxRows = 1048576;

    private static string ReadEntry(byte[] package, string path)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry(path)!.Open());
        return reader.ReadToEnd();
    }

    private static Spreadsheet CornerSheet()
    {
        var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "top-left");
        sheet.AddCell(MaxRows, MaxColumns, "bottom-right");
        return spreadsheet;
    }

    [Fact]
    public async Task Ods_CornerCells_GenerateQuickly()
    {
        using var spreadsheet = CornerSheet();

        var watch = Stopwatch.StartNew();
        var bytes = await spreadsheet.GenerateOdsFileAsync();
        watch.Stop();

        Assert.True(watch.ElapsedMilliseconds < 5000, $"Took {watch.ElapsedMilliseconds} ms");
        Assert.True(bytes.Length < 20_000, $"File is {bytes.Length} bytes");

        var content = ReadEntry(bytes, "content.xml");
        Assert.Contains($"table:number-rows-repeated=\"{MaxRows - 2}\"", content);
        Assert.Contains($"table:number-columns-repeated=\"{MaxColumns - 1}\"", content);
        Assert.True(SpreadsheetValidator.ValidateOds(bytes).IsValid);
    }

    [Fact]
    public async Task Xlsx_CornerCells_GenerateQuickly()
    {
        using var spreadsheet = CornerSheet();

        var watch = Stopwatch.StartNew();
        var bytes = await spreadsheet.GenerateXlsxFileAsync();
        watch.Stop();

        Assert.True(watch.ElapsedMilliseconds < 5000, $"Took {watch.ElapsedMilliseconds} ms");
        Assert.True(bytes.Length < 20_000, $"File is {bytes.Length} bytes");
        Assert.True(SpreadsheetValidator.ValidateXlsx(bytes).IsValid);
    }

    [Fact]
    public async Task Ods_CornerCells_RoundTrip()
    {
        using var source = CornerSheet();
        var bytes = await source.GenerateOdsFileAsync();

        using var spreadsheet = new Spreadsheet();
        var sheet = (await spreadsheet.ImportOdsFileAsync(bytes)).Sheets[0];

        Assert.Equal(2, sheet.Cells.Count);
        Assert.Equal("top-left", sheet.GetCell(1, 1)?.Value);
        Assert.Equal("bottom-right", sheet.GetCell(MaxRows, MaxColumns)?.Value);
    }

    /// <summary>
    /// Gaps of one and of several rows and columns, with typed and styled cells, keep their positions and values.
    /// </summary>
    [Fact]
    public async Task Ods_Gaps_RoundTrip()
    {
        using var source = new Spreadsheet();
        var sheet = source.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "a");
        sheet.AddCell(1, 3, "1.5", CellValueType.Float);
        sheet.AddCell(1, 7, "true", CellValueType.Boolean);
        sheet.AddCell(3, 2, "2024-01-31", CellValueType.DateTime);
        sheet.AddCell(10, 5, "b").Style = new CellStyle { Bold = true };
        sheet.AddCell(10, 6, "c").Formula = "=A1";
        var bytes = await source.GenerateOdsFileAsync();

        Assert.True(SpreadsheetValidator.ValidateOds(bytes).IsValid);

        using var spreadsheet = new Spreadsheet();
        var imported = (await spreadsheet.ImportOdsFileAsync(bytes)).Sheets[0];

        Assert.Equal(6, imported.Cells.Count);
        Assert.Equal("a", imported.GetCell(1, 1)?.Value);
        Assert.Equal(("1.5", CellValueType.Float), (imported.GetCell(1, 3)?.Value, imported.GetCell(1, 3)?.ValueType));
        Assert.Equal(("true", CellValueType.Boolean), (imported.GetCell(1, 7)?.Value, imported.GetCell(1, 7)?.ValueType));
        Assert.Equal(CellValueType.DateTime, imported.GetCell(3, 2)?.ValueType);
        Assert.True(imported.GetCell(10, 5)?.Style?.Bold);
        Assert.Equal("=A1", imported.GetCell(10, 6)?.Formula);
    }

    /// <summary>
    /// A row with a cell in the last column, and a cell in the last row, write no zero-count filler.
    /// </summary>
    [Fact]
    public async Task Ods_LastColumnAndRow_NoZeroRepeat()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, MaxColumns, "x");
        sheet.AddCell(MaxRows, 1, "y");

        var content = ReadEntry(await spreadsheet.GenerateOdsFileAsync(), "content.xml");

        Assert.DoesNotContain("repeated=\"0\"", content);
        Assert.DoesNotContain("repeated=\"-", content);
    }
}
