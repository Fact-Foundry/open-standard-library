using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class AutoFilterTests
{
    [Fact]
    public void AutoFilterRange_DefaultsToNull()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        Assert.Null(sheet.AutoFilterRange);
    }

    [Fact]
    public void SetAutoFilter_NoArgs_UsesFullRange()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Value");
        sheet.AddCell(2, 1, "A");
        sheet.AddCell(2, 2, "1");
        sheet.AddCell(3, 1, "B");
        sheet.AddCell(3, 2, "2");

        sheet.SetAutoFilter();

        Assert.NotNull(sheet.AutoFilterRange);
        Assert.Equal((1, 1, 3, 2), sheet.AutoFilterRange.Value);
    }

    [Fact]
    public void SetAutoFilter_ExplicitRange()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "Header");

        sheet.SetAutoFilter(1, 1, 10, 4);

        Assert.Equal((1, 1, 10, 4), sheet.AutoFilterRange.Value);
    }

    [Fact]
    public void SetAutoFilter_EmptySheet_DoesNothing()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.SetAutoFilter();

        Assert.Null(sheet.AutoFilterRange);
    }

    // --- XLSX ---

    [Fact]
    public async Task Xlsx_AutoFilter_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Age");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "30");
        sheet.AddCell(3, 1, "Bob");
        sheet.AddCell(3, 2, "25");
        sheet.SetAutoFilter();

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal("Name", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
        Assert.NotNull(workbook.Sheets[0].AutoFilterRange);
        Assert.Equal((1, 1, 3, 2), workbook.Sheets[0].AutoFilterRange.Value);
    }

    [Fact]
    public async Task Xlsx_AutoFilter_ExplicitRange_RoundTrips()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Header1");
        sheet.AddCell(1, 2, "Header2");
        sheet.AddCell(1, 3, "Header3");
        for (int r = 2; r <= 5; r++)
            for (int c = 1; c <= 3; c++)
                sheet.AddCell(r, c, $"R{r}C{c}");
        sheet.SetAutoFilter(1, 1, 5, 3);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal((1, 1, 5, 3), workbook.Sheets[0].AutoFilterRange.Value);
    }

    [Fact]
    public async Task Xlsx_NoAutoFilter_NoAutoFilterElement()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "No filter");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Null(workbook.Sheets[0].AutoFilterRange);
    }

    [Fact]
    public async Task Xlsx_AutoFilterWithStylingAndFreeze_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Full");
        sheet.FreezeRows = 1;

        var header = sheet.AddCell(1, 1, "Name");
        header.Style = new CellStyle { Bold = true, BackgroundColor = "#2196F3" };
        sheet.AddCell(1, 2, "Value");
        sheet.AddCell(2, 1, "Item");
        sheet.AddCell(2, 2, "100");
        sheet.SetAutoFilter();
        sheet.AutoFitColumns(minWidth: 10);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(1, workbook.Sheets[0].FreezeRows);
        Assert.NotNull(workbook.Sheets[0].AutoFilterRange);
        Assert.Equal("Name", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    // --- ODS ---

    [Fact]
    public async Task Ods_AutoFilter_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Age");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "30");
        sheet.AddCell(3, 1, "Bob");
        sheet.AddCell(3, 2, "25");
        sheet.SetAutoFilter();

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal("Name", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
        Assert.NotNull(workbook.Sheets[0].AutoFilterRange);
        Assert.Equal((1, 1, 3, 2), workbook.Sheets[0].AutoFilterRange.Value);
    }

    [Fact]
    public async Task Ods_AutoFilter_ExplicitRange_RoundTrips()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "H1");
        sheet.AddCell(1, 2, "H2");
        for (int r = 2; r <= 4; r++)
            for (int c = 1; c <= 2; c++)
                sheet.AddCell(r, c, $"R{r}C{c}");
        sheet.SetAutoFilter(1, 1, 4, 2);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal((1, 1, 4, 2), workbook.Sheets[0].AutoFilterRange.Value);
    }

    [Fact]
    public async Task Ods_NoAutoFilter_NoDbRange()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "No filter");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Null(workbook.Sheets[0].AutoFilterRange);
    }
}
