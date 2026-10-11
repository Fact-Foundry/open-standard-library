using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class CellTests
{
    /// <summary>
    /// Verifies a new cell has expected default values for all properties.
    /// </summary>
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

    /// <summary>
    /// Verifies the Value property can be set on a cell.
    /// </summary>
    [Fact]
    public void Cell_SetValue()
    {
        var cell = new oCell(2, 3);
        cell.Value = "Test";

        Assert.Equal("Test", cell.Value);
    }

    /// <summary>
    /// Verifies the Formula property can be set on a cell.
    /// </summary>
    [Fact]
    public void Cell_SetFormula()
    {
        var cell = new oCell(1, 1);
        cell.Formula = "=SUM(A1:A10)";

        Assert.Equal("=SUM(A1:A10)", cell.Formula);
    }

    /// <summary>
    /// Verifies the ValueType property can be set to Float.
    /// </summary>
    [Fact]
    public void Cell_SetValueType_Float()
    {
        var cell = new oCell(1, 1);
        cell.ValueType = CellValueType.Float;

        Assert.Equal(CellValueType.Float, cell.ValueType);
    }

    /// <summary>
    /// Verifies Row and Column are set at construction and are read-only.
    /// </summary>
    [Fact]
    public void Cell_RowAndColumn_AreReadOnly()
    {
        var cell = new oCell(5, 10);
        Assert.Equal(5, cell.Row);
        Assert.Equal(10, cell.Column);
    }

    /// <summary>
    /// Verifies the ValueType property can be set to Boolean while preserving the value.
    /// </summary>
    [Fact]
    public void Cell_SetValueType_Boolean()
    {
        var cell = new oCell(1, 1);
        cell.Value = "true";
        cell.ValueType = CellValueType.Boolean;

        Assert.Equal(CellValueType.Boolean, cell.ValueType);
        Assert.Equal("true", cell.Value);
    }

    /// <summary>
    /// Verifies the AsFloat extension method sets both the value and ValueType.
    /// </summary>
    [Fact]
    public void AsFloat_SetsValueAndType()
    {
        var cell = new oCell(1, 1);
        var result = cell.AsFloat(42.5);

        Assert.Equal("42.5", result.Value);
        Assert.Equal(CellValueType.Float, result.ValueType);
    }

    // --- Epoch conversion ---

    /// <summary>
    /// Verifies FromEpochSeconds converts a Unix timestamp to ISO 8601 DateTime.
    /// </summary>
    [Fact]
    public void FromEpochSeconds_ConvertsToDateTime()
    {
        var cell = new oCell(1, 1) { Value = "1747650600" }; // 2025-05-19T10:30:00 UTC
        cell.FromEpochSeconds();

        Assert.Equal(CellValueType.DateTime, cell.ValueType);
        Assert.Equal("2025-05-19T10:30:00", cell.Value);
    }

    /// <summary>
    /// Verifies FromEpochMilliseconds converts a Unix millisecond timestamp to ISO 8601 DateTime.
    /// </summary>
    [Fact]
    public void FromEpochMilliseconds_ConvertsToDateTime()
    {
        var cell = new oCell(1, 1) { Value = "1747650600000" }; // 2025-05-19T10:30:00 UTC
        cell.FromEpochMilliseconds();

        Assert.Equal(CellValueType.DateTime, cell.ValueType);
        Assert.Equal("2025-05-19T10:30:00", cell.Value);
    }

    /// <summary>
    /// Verifies ToEpochSeconds converts a DateTime value to Unix epoch seconds.
    /// </summary>
    [Fact]
    public void ToEpochSeconds_ConvertsFromDateTime()
    {
        var cell = new oCell(1, 1) { Value = "2025-05-19T10:30:00", ValueType = CellValueType.DateTime };
        cell.ToEpochSeconds();

        Assert.Equal(CellValueType.Int64, cell.ValueType);
        Assert.Equal("1747650600", cell.Value);
    }

    /// <summary>
    /// Verifies ToEpochMilliseconds converts a DateTime value to Unix epoch milliseconds.
    /// </summary>
    [Fact]
    public void ToEpochMilliseconds_ConvertsFromDateTime()
    {
        var cell = new oCell(1, 1) { Value = "2025-05-19T10:30:00", ValueType = CellValueType.DateTime };
        cell.ToEpochMilliseconds();

        Assert.Equal(CellValueType.Int64, cell.ValueType);
        Assert.Equal("1747650600000", cell.Value);
    }

    /// <summary>
    /// Verifies epoch round-trip: seconds → DateTime → seconds.
    /// </summary>
    [Fact]
    public void EpochSeconds_RoundTrip()
    {
        var cell = new oCell(1, 1) { Value = "1747650600" };
        cell.FromEpochSeconds();
        cell.ToEpochSeconds();

        Assert.Equal("1747650600", cell.Value);
    }

    /// <summary>
    /// Verifies FromEpochSeconds is a no-op for non-numeric values.
    /// </summary>
    [Fact]
    public void FromEpochSeconds_InvalidValue_NoChange()
    {
        var cell = new oCell(1, 1) { Value = "not-a-number" };
        cell.FromEpochSeconds();

        Assert.Equal("not-a-number", cell.Value);
        Assert.Equal(CellValueType.String, cell.ValueType);
    }
}
