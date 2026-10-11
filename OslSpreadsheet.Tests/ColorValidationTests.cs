using System.IO.Compression;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests that color values are validated before generation and normalized when written, so a color string can't
/// inject markup into a generated file.
/// </summary>
public class ColorValidationTests
{
    private static string ReadEntry(byte[] package, string path)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry(path)!.Open());
        return reader.ReadToEnd();
    }

    /// <summary>
    /// Every color property is checked, and the issue names the property and cell.
    /// </summary>
    [Theory]
    [InlineData("FontColor")]
    [InlineData("BackgroundColor")]
    [InlineData("BorderBottom.Color")]
    public void Validate_RejectsNonHexColor(string property)
    {
        var style = new CellStyle();
        const string bad = "#FF0000\"/><evil/><x a=\"";
        switch (property)
        {
            case "FontColor": style.FontColor = bad; break;
            case "BackgroundColor": style.BackgroundColor = bad; break;
            default: style.BorderBottom = new CellBorder { Color = bad }; break;
        }

        var workbook = new oWorkbook();
        workbook.AddSheet("Data").AddCell(2, 3, "x").Style = style;

        var result = workbook.Validate();

        Assert.False(result.IsValid);
        var issue = Assert.Single(result.Errors);
        Assert.Equal("WB_COLOR_INVALID", issue.Code);
        Assert.Contains(property, issue.Message);
        Assert.Equal("C2", issue.Cell);
    }

    /// <summary>
    /// Generation refuses the workbook, so the injected markup never reaches a file.
    /// </summary>
    [Fact]
    public async Task Generate_InvalidColor_Throws()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "x").Style = new CellStyle { BackgroundColor = "red" };

        await Assert.ThrowsAsync<InvalidWorkbookException>(() => spreadsheet.GenerateXlsxFileAsync());
        await Assert.ThrowsAsync<InvalidWorkbookException>(() => spreadsheet.GenerateOdsFileAsync());
    }

    /// <summary>
    /// Colors without a leading '#' and in lowercase are accepted and written normalized in both formats.
    /// </summary>
    [Fact]
    public async Task Generate_NormalizesAcceptedColors()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "x").Style = new CellStyle
        {
            FontColor = "ff0000",
            BackgroundColor = "#2196f3",
            BorderTop = new CellBorder { Color = "1565C0" }
        };

        Assert.True(spreadsheet.Workbook.Validate().IsValid);
        var xlsx = await spreadsheet.GenerateXlsxFileAsync();
        var ods = await spreadsheet.GenerateOdsFileAsync();

        var styles = ReadEntry(xlsx, "xl/styles.xml");
        Assert.Contains("<color rgb=\"FFFF0000\"/>", styles);
        Assert.Contains("<fgColor rgb=\"FF2196F3\"/>", styles);
        Assert.Contains("<color rgb=\"FF1565C0\"/>", styles);

        var content = ReadEntry(ods, "content.xml");
        Assert.Contains("fo:color=\"#FF0000\"", content);
        Assert.Contains("fo:background-color=\"#2196F3\"", content);
        Assert.Contains("solid #1565C0", content);

        Assert.Empty(SpreadsheetValidator.ValidateXlsx(xlsx).Issues);
        Assert.Empty(SpreadsheetValidator.ValidateOds(ods).Issues);
    }

    /// <summary>
    /// Values that look close to a color but aren't (wrong length, non-hex digits) are rejected.
    /// </summary>
    [Theory]
    [InlineData("#FFF")]
    [InlineData("#GG0000")]
    [InlineData("#FF00000")]
    public void Validate_RejectsNearMisses(string color)
    {
        var workbook = new oWorkbook();
        workbook.AddSheet("Data").AddCell(1, 1, "x").Style = new CellStyle { FontColor = color };

        Assert.Equal("WB_COLOR_INVALID", Assert.Single(workbook.Validate().Errors).Code);
    }
}
