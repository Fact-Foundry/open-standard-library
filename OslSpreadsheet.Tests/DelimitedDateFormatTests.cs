using System.Globalization;
using System.Text;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for <see cref="oWorkbook.DelimitedDateFormat"/>, which re-formats DateTime cells on delimited export.
/// </summary>
public class DelimitedDateFormatTests
{
    private static async Task<string> GenerateCsvAsync(string? format, params (string Value, CellValueType Type)[] cells)
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.DelimitedDateFormat = format;
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        for (int i = 0; i < cells.Length; i++)
            sheet.AddCell(1, i + 1, cells[i].Value, cells[i].Type);
        return Encoding.UTF8.GetString(await spreadsheet.GenerateCsvFileAsync()).TrimEnd();
    }

    [Theory]
    [InlineData("yyyy-MM-dd", "2024-01-31T13:45:00", "2024-01-31")]
    [InlineData("MM/dd/yyyy", "2024-01-31", "01/31/2024")]
    [InlineData("dd MMM yyyy HH:mm", "2024-01-31T13:45:00", "31 Jan 2024 13:45")]
    [InlineData("yyyy-MM-dd", "1/31/2024", "2024-01-31")]
    public async Task Format_IsAppliedToDateTimeCells(string format, string value, string expected)
    {
        var csv = await GenerateCsvAsync(format, (value, CellValueType.DateTime));

        Assert.Equal($"\"{expected}\"", csv);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public async Task NoFormat_WritesValueAsStored(string? format)
    {
        var csv = await GenerateCsvAsync(format, ("2024-01-31T13:45:00", CellValueType.DateTime));

        Assert.Equal("\"2024-01-31T13:45:00\"", csv);
    }

    [Fact]
    public async Task Format_LeavesOtherCellTypesAlone()
    {
        var csv = await GenerateCsvAsync("yyyy-MM-dd",
            ("2024-01-31T13:45:00", CellValueType.String),
            ("45322", CellValueType.Float),
            ("2024-01-31T13:45:00", CellValueType.DateTime));

        Assert.Equal("\"2024-01-31T13:45:00\",\"45322\",\"2024-01-31\"", csv);
    }

    [Fact]
    public async Task Format_UsesInvariantCulture()
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");

            var csv = await GenerateCsvAsync("MMM d, yyyy", ("2024-03-05", CellValueType.DateTime));

            Assert.Equal("\"Mar 5, 2024\"", csv);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public async Task Format_AppliesToImportedXlsxDates()
    {
        using var source = new Spreadsheet();
        source.Workbook.AddSheet("Data").AddCell(1, 1, "2024-01-31", CellValueType.DateTime);
        var xlsx = await source.GenerateXlsxFileAsync();

        using var spreadsheet = new Spreadsheet();
        var workbook = await spreadsheet.ImportXlsxFileAsync(xlsx);
        workbook.DelimitedDateFormat = "yyyy-MM-dd";
        var csv = Encoding.UTF8.GetString(await spreadsheet.GenerateCsvFileAsync()).TrimEnd();

        Assert.Equal("\"2024-01-31\"", csv);
    }

    [Fact]
    public void InvalidFormat_IsValidationError()
    {
        var workbook = new oWorkbook { DelimitedDateFormat = "%" };
        workbook.AddSheet("Data").AddCell(1, 1, "2024-01-31", CellValueType.DateTime);

        var result = workbook.Validate(ValidationFileFormat.Delimited);

        Assert.Contains(result.Issues, i => i.Code == "WB_DELIMITED_DATE_FORMAT_INVALID" && i.Severity == ValidationSeverity.Error);
    }

    [Fact]
    public void InvalidFormat_IgnoredForOtherFormats()
    {
        var workbook = new oWorkbook { DelimitedDateFormat = "%" };
        workbook.AddSheet("Data").AddCell(1, 1, "2024-01-31", CellValueType.DateTime);

        var result = workbook.Validate(ValidationFileFormat.Xlsx);

        Assert.DoesNotContain(result.Issues, i => i.Code == "WB_DELIMITED_DATE_FORMAT_INVALID");
    }

    [Fact]
    public async Task InvalidFormat_GenerateThrowsInvalidWorkbookException()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.DelimitedDateFormat = "%";
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "2024-01-31", CellValueType.DateTime);

        await Assert.ThrowsAsync<InvalidWorkbookException>(() => spreadsheet.GenerateCsvFileAsync());
    }
}
