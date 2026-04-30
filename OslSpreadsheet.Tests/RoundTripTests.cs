using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class RoundTripTests
{
    private static oWorkbook BuildTestWorkbook(oWorkbook workbook)
    {
        var sheet = workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Score");
        sheet.AddCell(1, 3, "Notes");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "95.5", CellValueType.Float);
        sheet.AddCell(2, 3, "Top performer");
        sheet.AddCell(3, 1, "Bob");
        sheet.AddCell(3, 2, "82.0", CellValueType.Float);
        sheet.AddCell(3, 3, "");
        return workbook;
    }

    // --- ODS ---

    [Fact]
    public async Task Ods_RoundTrip_PreservesSheetName()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Single(workbook.Sheets);
        Assert.Equal("Data", workbook.Sheets[0].SheetName);
    }

    [Fact]
    public async Task Ods_RoundTrip_PreservesCellValues()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        Assert.Equal("Name", sheet.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Alice", sheet.GetRow(2).First(c => c.Column == 1).Value);
        Assert.Equal("Bob", sheet.GetRow(3).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Ods_RoundTrip_PreservesFloatValues()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        var scoreCell = sheet.GetRow(2).First(c => c.Column == 2);
        Assert.Equal("95.5", scoreCell.Value);
        Assert.Equal(CellValueType.Float, scoreCell.ValueType);
    }

    [Fact]
    public async Task Ods_RoundTrip_PreservesRowAndColumnCounts()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        Assert.Equal(3, sheet.RowCount);
        Assert.Equal(3, sheet.ColumnCount);
    }

    [Fact]
    public async Task Ods_RoundTrip_MultipleSheets()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet1 = spreadsheet.Workbook.AddSheet("First");
        sheet1.AddCell(1, 1, "Sheet1Data");
        var sheet2 = spreadsheet.Workbook.AddSheet("Second");
        sheet2.AddCell(1, 1, "Sheet2Data");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal(2, workbook.Sheets.Count);
        Assert.Equal("First", workbook.Sheets[0].SheetName);
        Assert.Equal("Second", workbook.Sheets[1].SheetName);
        Assert.Equal("Sheet1Data", workbook.Sheets[0].Cells[0].Value);
        Assert.Equal("Sheet2Data", workbook.Sheets[1].Cells[0].Value);
    }

    [Fact]
    public async Task Ods_Generate_ProducesNonEmptyBytes()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    // --- XLSX ---

    [Fact]
    public async Task Xlsx_RoundTrip_PreservesSheetName()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Single(workbook.Sheets);
        Assert.Equal("Data", workbook.Sheets[0].SheetName);
    }

    [Fact]
    public async Task Xlsx_RoundTrip_PreservesCellValues()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        Assert.Equal("Name", sheet.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Alice", sheet.GetRow(2).First(c => c.Column == 1).Value);
        Assert.Equal("Bob", sheet.GetRow(3).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Xlsx_RoundTrip_PreservesFloatValues()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        var scoreCell = sheet.GetRow(2).First(c => c.Column == 2);
        Assert.Equal("95.5", scoreCell.Value);
        Assert.Equal(CellValueType.Float, scoreCell.ValueType);
    }

    [Fact]
    public async Task Xlsx_RoundTrip_PreservesRowAndColumnCounts()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        Assert.Equal(3, sheet.RowCount);
        Assert.Equal(3, sheet.ColumnCount);
    }

    [Fact]
    public async Task Xlsx_RoundTrip_MultipleSheets()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet1 = spreadsheet.Workbook.AddSheet("First");
        sheet1.AddCell(1, 1, "Sheet1Data");
        var sheet2 = spreadsheet.Workbook.AddSheet("Second");
        sheet2.AddCell(1, 1, "Sheet2Data");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(2, workbook.Sheets.Count);
        Assert.Equal("First", workbook.Sheets[0].SheetName);
        Assert.Equal("Second", workbook.Sheets[1].SheetName);
        Assert.Equal("Sheet1Data", workbook.Sheets[0].Cells[0].Value);
        Assert.Equal("Sheet2Data", workbook.Sheets[1].Cells[0].Value);
    }

    [Fact]
    public async Task Xlsx_Generate_ProducesNonEmptyBytes()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    // --- CSV ---

    [Fact]
    public async Task Csv_RoundTrip_PreservesCellValues()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Score");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "95.5");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportCsvFileAsync(bytes);
        var imported = workbook.Sheets[0];

        Assert.Equal("Name", imported.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Score", imported.GetRow(1).First(c => c.Column == 2).Value);
        Assert.Equal("Alice", imported.GetRow(2).First(c => c.Column == 1).Value);
        Assert.Equal("95.5", imported.GetRow(2).First(c => c.Column == 2).Value);
    }

    [Fact]
    public async Task Csv_Generate_ProducesNonEmptyBytes()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Hello");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public async Task Csv_RoundTrip_HandlesCommasInValues()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Value, with comma");
        sheet.AddCell(1, 2, "Normal");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportCsvFileAsync(bytes);
        var imported = workbook.Sheets[0];

        Assert.Equal("Value, with comma", imported.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Normal", imported.GetRow(1).First(c => c.Column == 2).Value);
    }

    // CSV import does not unescape doubled quotes (e.g. 5"" Fitting -> 5"" Fitting instead of 5" Fitting)
    [Fact]
    public async Task Csv_RoundTrip_EmbeddedQuotes_KnownBug()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "5\" Fitting");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportCsvFileAsync(bytes);
        var imported = workbook.Sheets[0];

        // Current behavior: doubled quotes are NOT unescaped on import
        Assert.Equal("5\"\" Fitting", imported.GetRow(1).First(c => c.Column == 1).Value);
    }

    // --- Empty workbook ---

    [Fact]
    public async Task Ods_EmptySheet_GeneratesWithoutError()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Empty");

        var bytes = await spreadsheet.GenerateOdsFileAsync();
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public async Task Xlsx_EmptySheet_GeneratesWithoutError()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Empty");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();
        Assert.True(bytes.Length > 0);
    }
}
