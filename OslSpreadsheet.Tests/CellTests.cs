using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class CellTests
{
    [Fact]
    public void Cell_DefaultValues()
    {
        var cell = new oCell(1, 1);

        Assert.Equal(1, cell.Row);
        Assert.Equal(1, cell.Column);
        Assert.Equal("", cell.Value);
        Assert.Equal(CellValueType.String, cell.ValueType);
        Assert.Null(cell.Formula);
    }

    [Fact]
    public void Cell_SetValue()
    {
        var cell = new oCell(2, 3);
        cell.Value = "Test";

        Assert.Equal("Test", cell.Value);
    }

    [Fact]
    public void Cell_SetFormula()
    {
        var cell = new oCell(1, 1);
        cell.Formula = "=SUM(A1:A10)";

        Assert.Equal("=SUM(A1:A10)", cell.Formula);
    }

    [Fact]
    public void Cell_SetValueType_Float()
    {
        var cell = new oCell(1, 1);
        cell.ValueType = CellValueType.Float;

        Assert.Equal(CellValueType.Float, cell.ValueType);
    }

    [Fact]
    public void Cell_RowAndColumn_AreReadOnly()
    {
        var cell = new oCell(5, 10);
        Assert.Equal(5, cell.Row);
        Assert.Equal(10, cell.Column);
    }

    [Fact]
    public void AsFloat_SetsValueAndType()
    {
        var cell = new oCell(1, 1);
        var result = cell.AsFloat<float>(42.5f);

        Assert.Equal("42.5", result.Value);
        Assert.Equal(CellValueType.Float, result.ValueType);
    }
}
