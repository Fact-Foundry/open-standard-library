using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests that cell styles and column widths are read back on import from XLSX and ODS.
/// </summary>
public class StyleImportTests
{
    private static CellStyle FullStyle() => new()
    {
        Bold = true,
        Italic = true,
        Underline = true,
        FontColor = "#FF0000",
        BackgroundColor = "#FFFF00",
        FontName = "Arial",
        FontSize = 14,
        WrapText = true,
        NumberFormat = "#,##0.00",
        BorderTop = new CellBorder { Style = BorderStyle.Thin },
        BorderBottom = new CellBorder { Style = BorderStyle.Medium, Color = "#1565C0" },
        BorderLeft = new CellBorder { Style = BorderStyle.Thick, Color = "#000000" }
    };

    private static void AssertStyleEqual(CellStyle expected, CellStyle? actual)
    {
        Assert.NotNull(actual);
        Assert.Equal(expected.Bold, actual.Bold);
        Assert.Equal(expected.Italic, actual.Italic);
        Assert.Equal(expected.Underline, actual.Underline);
        Assert.Equal(expected.FontColor, actual.FontColor);
        Assert.Equal(expected.BackgroundColor, actual.BackgroundColor);
        Assert.Equal(expected.FontName, actual.FontName);
        Assert.Equal(expected.FontSize, actual.FontSize);
        Assert.Equal(expected.WrapText, actual.WrapText);
        Assert.Equal(expected.NumberFormat, actual.NumberFormat);
        AssertBorderEqual(expected.BorderTop, actual.BorderTop);
        AssertBorderEqual(expected.BorderBottom, actual.BorderBottom);
        AssertBorderEqual(expected.BorderLeft, actual.BorderLeft);
        AssertBorderEqual(expected.BorderRight, actual.BorderRight);
    }

    private static void AssertBorderEqual(CellBorder? expected, CellBorder? actual)
    {
        if (expected == null) { Assert.Null(actual); return; }
        Assert.NotNull(actual);
        Assert.Equal(expected.Style, actual.Style);
        Assert.Equal(expected.Color, actual.Color);
    }

    private static async Task<oSpreadsheet> RoundTripAsync(Spreadsheet spreadsheet, bool ods)
    {
        var bytes = ods ? await spreadsheet.GenerateOdsFileAsync() : await spreadsheet.GenerateXlsxFileAsync();
        using var importer = new Spreadsheet();
        var workbook = ods ? await importer.ImportOdsFileAsync(bytes) : await importer.ImportXlsxFileAsync(bytes);
        return workbook.Sheets[0];
    }

    /// <summary>
    /// Every CellStyle property survives a round-trip through both formats. ODS always stores a border color,
    /// so a border written without one comes back as black from ODS.
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RoundTrip_PreservesFullStyle(bool ods)
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "1234.5", CellValueType.Float).Style = FullStyle();
        sheet.AddCell(1, 2, "plain");

        var imported = await RoundTripAsync(spreadsheet, ods);

        var expected = FullStyle();
        if (ods) expected.BorderTop!.Color = "#000000";
        AssertStyleEqual(expected, imported.GetCell(1, 1)?.Style);
        Assert.Null(imported.GetCell(1, 2)?.Style);
    }

    /// <summary>
    /// Partial styles come back with only the properties that were set; defaults such as the font name stay null.
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RoundTrip_PartialStyle_KeepsUnsetPropertiesNull(bool ods)
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Header").Style = new CellStyle { Bold = true, BackgroundColor = "#2196F3" };

        var style = (await RoundTripAsync(spreadsheet, ods)).GetCell(1, 1)?.Style;

        AssertStyleEqual(new CellStyle { Bold = true, BackgroundColor = "#2196F3" }, style);
    }

    /// <summary>
    /// Cells sharing a style on export get independent Style objects on import.
    /// </summary>
    [Fact]
    public async Task Import_CellsGetIndependentStyleObjects()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        var shared = new CellStyle { Bold = true };
        sheet.AddCell(1, 1, "a").Style = shared;
        sheet.AddCell(1, 2, "b").Style = shared;

        var imported = await RoundTripAsync(spreadsheet, ods: false);
        imported.GetCell(1, 1)!.Style!.Italic = true;

        Assert.False(imported.GetCell(1, 2)!.Style!.Italic);
    }

    /// <summary>
    /// Explicit column widths survive a round-trip through both formats, and sheets without widths stay without them.
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RoundTrip_PreservesColumnWidths(bool ods)
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "a");
        sheet.AddCell(1, 3, "c");
        sheet.SetColumnWidth(1, 25);
        sheet.SetColumnWidth(3, 5.5);
        spreadsheet.Workbook.AddSheet("Plain").AddCell(1, 1, "x");

        var bytes = ods ? await spreadsheet.GenerateOdsFileAsync() : await spreadsheet.GenerateXlsxFileAsync();
        using var importer = new Spreadsheet();
        var workbook = ods ? await importer.ImportOdsFileAsync(bytes) : await importer.ImportXlsxFileAsync(bytes);

        var widths = workbook.Sheets[0].ColumnWidths;
        Assert.Equal(25, widths[1], 1);
        Assert.Equal(5.5, widths[3], 1);
        Assert.False(widths.ContainsKey(2));
        Assert.Empty(workbook.Sheets[1].ColumnWidths);
    }

    /// <summary>
    /// DateTime cells with the default date format and no other styling still import with Style null.
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RoundTrip_DefaultDateCell_HasNoStyle(bool ods)
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "2024-01-31T13:45:00", CellValueType.DateTime);

        var cell = (await RoundTripAsync(spreadsheet, ods)).GetCell(1, 1);

        Assert.Equal(CellValueType.DateTime, cell?.ValueType);
        Assert.Null(cell?.Style);
    }
}
