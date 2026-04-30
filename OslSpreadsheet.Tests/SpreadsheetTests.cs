using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class SpreadsheetTests
{
    [Fact]
    public void NewSheet_HasNoCells()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        Assert.Empty(sheet.Cells);
    }

    [Fact]
    public void AddCell_EmptyCell_AddsWithDefaults()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        var cell = sheet.AddCell(1, 1);

        Assert.Single(sheet.Cells);
        Assert.Equal(1, cell.Row);
        Assert.Equal(1, cell.Column);
        Assert.Equal("", cell.Value);
        Assert.Equal(CellValueType.String, cell.ValueType);
    }

    [Fact]
    public void AddCell_WithValue_SetsValue()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        var cell = sheet.AddCell(1, 1, "Hello");

        Assert.Equal("Hello", cell.Value);
        Assert.Equal(CellValueType.String, cell.ValueType);
    }

    [Fact]
    public void AddCell_WithValueType_SetsValueType()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        var cell = sheet.AddCell(1, 1, "42.5", CellValueType.Float);

        Assert.Equal("42.5", cell.Value);
        Assert.Equal(CellValueType.Float, cell.ValueType);
    }

    [Fact]
    public void AddCell_SamePosition_ReplacesExistingCell()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "First");
        sheet.AddCell(1, 1, "Second");

        Assert.Single(sheet.Cells);
        Assert.Equal("Second", sheet.Cells[0].Value);
    }

    [Fact]
    public void RowCount_EmptySheet_ReturnsZero()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        Assert.Equal(0, sheet.RowCount);
    }

    [Fact]
    public void ColumnCount_EmptySheet_ReturnsZero()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        Assert.Equal(0, sheet.ColumnCount);
    }

    [Fact]
    public void RowCount_ReturnsMaxRowIndex()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "A");
        sheet.AddCell(3, 1, "B");
        sheet.AddCell(5, 2, "C");

        Assert.Equal(5, sheet.RowCount);
    }

    [Fact]
    public void ColumnCount_ReturnsMaxColumnIndex()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "A");
        sheet.AddCell(1, 4, "B");
        sheet.AddCell(2, 2, "C");

        Assert.Equal(4, sheet.ColumnCount);
    }

    [Fact]
    public void GetRow_ReturnsCellsInColumnOrder()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 3, "C");
        sheet.AddCell(1, 1, "A");
        sheet.AddCell(1, 2, "B");
        sheet.AddCell(2, 1, "Other row");

        var row = sheet.GetRow(1);

        Assert.Equal(3, row.Count);
        Assert.Equal("A", row[0].Value);
        Assert.Equal("B", row[1].Value);
        Assert.Equal("C", row[2].Value);
    }

    [Fact]
    public void GetRow_NonexistentRow_ReturnsEmpty()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "A");

        var row = sheet.GetRow(99);
        Assert.Empty(row);
    }

    [Fact]
    public void ToArray_ReturnsCorrect2DArray()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "A1");
        sheet.AddCell(1, 2, "B1");
        sheet.AddCell(2, 1, "A2");
        sheet.AddCell(2, 2, "B2");

        var array = sheet.ToArray();

        Assert.Equal(2, array.GetLength(0));
        Assert.Equal(2, array.GetLength(1));
        Assert.Equal("A1", array[0, 0]);
        Assert.Equal("B1", array[0, 1]);
        Assert.Equal("A2", array[1, 0]);
        Assert.Equal("B2", array[1, 1]);
    }

    [Fact]
    public void ToArray_SparseData_HasNulls()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        sheet.AddCell(1, 1, "A1");
        sheet.AddCell(2, 2, "B2");

        var array = sheet.ToArray();

        Assert.Equal("A1", array[0, 0]);
        Assert.Null(array[0, 1]);
        Assert.Null(array[1, 0]);
        Assert.Equal("B2", array[1, 1]);
    }

    [Fact]
    public async Task AddCellAsync_WithValue_SetsValue()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        var cell = await sheet.AddCellAsync(1, 1, "Async Value");

        Assert.Equal("Async Value", cell.Value);
    }

    [Fact]
    public async Task AddCellAsync_WithValueType_SetsValueType()
    {
        var sheet = new oSpreadsheet(1, "Sheet1");
        var cell = await sheet.AddCellAsync(1, 1, "99.9", CellValueType.Float);

        Assert.Equal("99.9", cell.Value);
        Assert.Equal(CellValueType.Float, cell.ValueType);
    }

    [Fact]
    public void SheetName_CanBeChanged()
    {
        var sheet = new oSpreadsheet(1, "Original");
        sheet.SheetName = "Renamed";
        Assert.Equal("Renamed", sheet.SheetName);
    }

    [Fact]
    public void Index_IsReadOnly()
    {
        var sheet = new oSpreadsheet(7, "Sheet7");
        Assert.Equal(7, sheet.Index);
    }
}
