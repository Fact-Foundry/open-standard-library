using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class ColumnWidthTests
{
    [Fact]
    public void ColumnWidths_DefaultsToEmpty()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        Assert.Empty(sheet.ColumnWidths);
    }

    [Fact]
    public void SetColumnWidth_StoresWidth()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.SetColumnWidth(1, 20);

        Assert.Single(sheet.ColumnWidths);
        Assert.Equal(20, sheet.ColumnWidths[1]);
    }

    [Fact]
    public void SetColumnWidth_MultipleColumns()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.SetColumnWidth(1, 10);
        sheet.SetColumnWidth(2, 30);
        sheet.SetColumnWidth(3, 15);

        Assert.Equal(3, sheet.ColumnWidths.Count);
        Assert.Equal(10, sheet.ColumnWidths[1]);
        Assert.Equal(30, sheet.ColumnWidths[2]);
        Assert.Equal(15, sheet.ColumnWidths[3]);
    }

    [Fact]
    public void SetColumnWidth_OverwritesExisting()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.SetColumnWidth(1, 10);
        sheet.SetColumnWidth(1, 25);

        Assert.Single(sheet.ColumnWidths);
        Assert.Equal(25, sheet.ColumnWidths[1]);
    }

    [Fact]
    public void AutoFitColumns_SetsWidthsBasedOnContent()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "Short");
        sheet.AddCell(1, 2, "A much longer string value");
        sheet.AddCell(2, 1, "Also short");
        sheet.AddCell(2, 2, "Tiny");

        sheet.AutoFitColumns();

        Assert.Equal(2, sheet.ColumnWidths.Count);
        Assert.Equal(10 + 2, sheet.ColumnWidths[1]);
        Assert.Equal(26 + 2, sheet.ColumnWidths[2]);
    }

    [Fact]
    public void AutoFitColumns_RespectsMinWidth()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "Hi");

        sheet.AutoFitColumns(minWidth: 15);

        Assert.Equal(15, sheet.ColumnWidths[1]);
    }

    [Fact]
    public void AutoFitColumns_RespectsMaxWidth()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, new string('x', 200));

        sheet.AutoFitColumns(maxWidth: 60);

        Assert.Equal(60, sheet.ColumnWidths[1]);
    }

    [Fact]
    public void AutoFitColumns_EmptySheet_DoesNothing()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AutoFitColumns();

        Assert.Empty(sheet.ColumnWidths);
    }

    // --- XLSX ---

    [Fact]
    public async Task Xlsx_ColumnWidths_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Description");
        sheet.SetColumnWidth(1, 15);
        sheet.SetColumnWidth(2, 40);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal("Name", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Description", workbook.Sheets[0].GetRow(1).First(c => c.Column == 2).Value);
    }

    [Fact]
    public async Task Xlsx_AutoFitColumns_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "ID");
        sheet.AddCell(1, 2, "A very long description column header");
        sheet.AddCell(2, 1, "1");
        sheet.AddCell(2, 2, "Short");
        sheet.AutoFitColumns(minWidth: 10, maxWidth: 60);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(2, workbook.Sheets[0].RowCount);
        Assert.Equal(2, workbook.Sheets[0].ColumnCount);
    }

    [Fact]
    public async Task Xlsx_NoColumnWidths_NoColsElement()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "No widths");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal("No widths", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Xlsx_ColumnWidthsWithStylingAndFreeze_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Full Featured");
        sheet.FreezeRows = 1;

        var header = sheet.AddCell(1, 1, "Name");
        header.Style = new CellStyle { Bold = true, BackgroundColor = "#2196F3" };
        sheet.AddCell(1, 2, "Description");
        sheet.AddCell(2, 1, "Item1");
        sheet.AddCell(2, 2, "A long description");

        sheet.AutoFitColumns(minWidth: 10, maxWidth: 60);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(1, workbook.Sheets[0].FreezeRows);
        Assert.Equal("Name", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    // --- ODS ---

    [Fact]
    public async Task Ods_ColumnWidths_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Description");
        sheet.SetColumnWidth(1, 15);
        sheet.SetColumnWidth(2, 40);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal("Name", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Description", workbook.Sheets[0].GetRow(1).First(c => c.Column == 2).Value);
    }

    [Fact]
    public async Task Ods_AutoFitColumns_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "ID");
        sheet.AddCell(1, 2, "A very long description column header");
        sheet.AddCell(2, 1, "1");
        sheet.AddCell(2, 2, "Short");
        sheet.AutoFitColumns(minWidth: 10, maxWidth: 60);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal(2, workbook.Sheets[0].RowCount);
        Assert.Equal(2, workbook.Sheets[0].ColumnCount);
    }

    [Fact]
    public async Task Ods_NoColumnWidths_DefaultBehavior()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "No widths");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal("No widths", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }
}
