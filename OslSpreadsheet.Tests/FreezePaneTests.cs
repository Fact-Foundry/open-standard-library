using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class FreezePaneTests
{
    [Fact]
    public void FreezeRows_DefaultsToZero()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        Assert.Equal(0, sheet.FreezeRows);
    }

    [Fact]
    public void FreezeColumns_DefaultsToZero()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        Assert.Equal(0, sheet.FreezeColumns);
    }

    [Fact]
    public void FreezeRows_CanBeSet()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.FreezeRows = 1;
        Assert.Equal(1, sheet.FreezeRows);
    }

    [Fact]
    public void FreezeColumns_CanBeSet()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.FreezeColumns = 2;
        Assert.Equal(2, sheet.FreezeColumns);
    }

    // --- XLSX ---

    [Fact]
    public async Task Xlsx_FreezeRows_RoundTrip()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.FreezeRows = 1;
        sheet.AddCell(1, 1, "Header");
        sheet.AddCell(2, 1, "Data");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(1, workbook.Sheets[0].FreezeRows);
        Assert.Equal(0, workbook.Sheets[0].FreezeColumns);
    }

    [Fact]
    public async Task Xlsx_FreezeColumns_RoundTrip()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.FreezeColumns = 2;
        sheet.AddCell(1, 1, "Col1");
        sheet.AddCell(1, 2, "Col2");
        sheet.AddCell(1, 3, "Col3");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(0, workbook.Sheets[0].FreezeRows);
        Assert.Equal(2, workbook.Sheets[0].FreezeColumns);
    }

    [Fact]
    public async Task Xlsx_FreezeRowsAndColumns_RoundTrip()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.FreezeRows = 1;
        sheet.FreezeColumns = 1;
        sheet.AddCell(1, 1, "Corner");
        sheet.AddCell(1, 2, "Header");
        sheet.AddCell(2, 1, "Row");
        sheet.AddCell(2, 2, "Data");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(1, workbook.Sheets[0].FreezeRows);
        Assert.Equal(1, workbook.Sheets[0].FreezeColumns);
    }

    [Fact]
    public async Task Xlsx_NoFreeze_NoFreezeOnImport()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "No freeze");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(0, workbook.Sheets[0].FreezeRows);
        Assert.Equal(0, workbook.Sheets[0].FreezeColumns);
    }

    [Fact]
    public async Task Xlsx_FreezeWithStyling_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Styled Frozen");
        sheet.FreezeRows = 1;

        var header = sheet.AddCell(1, 1, "Header");
        header.Style = new CellStyle { Bold = true, BackgroundColor = "#2196F3" };
        sheet.AddCell(2, 1, "Data");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(1, workbook.Sheets[0].FreezeRows);
        Assert.Equal("Header", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    // --- ODS ---

    [Fact]
    public async Task Ods_FreezeRows_RoundTrip()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.FreezeRows = 1;
        sheet.AddCell(1, 1, "Header");
        sheet.AddCell(2, 1, "Data");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal(1, workbook.Sheets[0].FreezeRows);
        Assert.Equal(0, workbook.Sheets[0].FreezeColumns);
    }

    [Fact]
    public async Task Ods_FreezeColumns_RoundTrip()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.FreezeColumns = 2;
        sheet.AddCell(1, 1, "Col1");
        sheet.AddCell(1, 2, "Col2");
        sheet.AddCell(1, 3, "Col3");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal(0, workbook.Sheets[0].FreezeRows);
        Assert.Equal(2, workbook.Sheets[0].FreezeColumns);
    }

    [Fact]
    public async Task Ods_FreezeRowsAndColumns_RoundTrip()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.FreezeRows = 1;
        sheet.FreezeColumns = 1;
        sheet.AddCell(1, 1, "Corner");
        sheet.AddCell(2, 2, "Data");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal(1, workbook.Sheets[0].FreezeRows);
        Assert.Equal(1, workbook.Sheets[0].FreezeColumns);
    }

    [Fact]
    public async Task Ods_NoFreeze_NoSettingsFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "No freeze");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal(0, workbook.Sheets[0].FreezeRows);
        Assert.Equal(0, workbook.Sheets[0].FreezeColumns);
    }
}
