using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class WorkbookTests
{
    [Fact]
    public void NewWorkbook_HasNoSheets()
    {
        var workbook = new oWorkbook();
        Assert.Empty(workbook.Sheets);
    }

    [Fact]
    public void AddSheet_WithoutName_UsesDefaultName()
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet();
        Assert.Equal("Sheet1", sheet.SheetName);
    }

    [Fact]
    public void AddSheet_WithName_UsesProvidedName()
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet("Sales");
        Assert.Equal("Sales", sheet.SheetName);
    }

    [Fact]
    public void AddSheet_MultipleTimes_IncrementsIndex()
    {
        var workbook = new oWorkbook();
        var sheet1 = workbook.AddSheet();
        var sheet2 = workbook.AddSheet();
        var sheet3 = workbook.AddSheet();

        Assert.Equal(1, sheet1.Index);
        Assert.Equal(2, sheet2.Index);
        Assert.Equal(3, sheet3.Index);
        Assert.Equal(3, workbook.Sheets.Count);
    }

    [Fact]
    public void AddSheet_MultipleTimes_DefaultNamesIncrement()
    {
        var workbook = new oWorkbook();
        workbook.AddSheet();
        workbook.AddSheet();
        workbook.AddSheet();

        Assert.Equal("Sheet1", workbook.Sheets[0].SheetName);
        Assert.Equal("Sheet2", workbook.Sheets[1].SheetName);
        Assert.Equal("Sheet3", workbook.Sheets[2].SheetName);
    }

    [Fact]
    public async Task AddSheetAsync_WithoutName_UsesDefaultName()
    {
        var workbook = new oWorkbook();
        var sheet = await workbook.AddSheetAsync();
        Assert.Equal("Sheet1", sheet.SheetName);
    }

    [Fact]
    public async Task AddSheetAsync_WithName_UsesProvidedName()
    {
        var workbook = new oWorkbook();
        var sheet = await workbook.AddSheetAsync("Async Sheet");
        Assert.Equal("Async Sheet", sheet.SheetName);
    }

    [Fact]
    public void Workbook_DefaultProperties()
    {
        var workbook = new oWorkbook();
        Assert.Equal("Open Standard Library v1", workbook.Generator);
        Assert.Equal("", workbook.InitialCreator);
        Assert.Equal("", workbook.Creator);
        Assert.Equal("", workbook.CreationDate);
        Assert.Equal(ColumnDelimeter.Comma, workbook.ColumnDelimeter);
    }

    [Fact]
    public void Workbook_SetCreator()
    {
        var workbook = new oWorkbook();
        workbook.Creator = "Test User";
        Assert.Equal("Test User", workbook.Creator);
    }
}
