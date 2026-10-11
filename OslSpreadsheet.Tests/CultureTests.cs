using System.Globalization;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests that generation and import produce the same results regardless of the thread's current culture.
/// File formats use '.' as the decimal separator; a culture using ',' must not leak into the XML.
/// </summary>
public class CultureTests
{
    /// <summary>
    /// Runs an action with the current thread culture set to one that uses a comma decimal separator.
    /// </summary>
    private static async Task WithCommaDecimalCultureAsync(Func<Task> action)
    {
        var original = (CultureInfo.CurrentCulture, CultureInfo.CurrentUICulture);
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
            CultureInfo.CurrentUICulture = CultureInfo.CurrentCulture;
            Assert.Equal("1,5", 1.5.ToString()); // sanity check that the culture is in effect
            await action();
        }
        finally
        {
            (CultureInfo.CurrentCulture, CultureInfo.CurrentUICulture) = original;
        }
    }

    private static Spreadsheet BuildWorkbook()
    {
        var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "2024-01-31T12:30:00", CellValueType.DateTime);
        sheet.AddCell(1, 2, "x").Style = new CellStyle { FontSize = 14.5 };
        sheet.AddCell(1, 3, "1234.5", CellValueType.Float);
        sheet.SetColumnWidth(1, 20.5);
        return spreadsheet;
    }

    /// <summary>
    /// Generated XLSX and ODS validate cleanly and contain period decimals under a comma-decimal culture.
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Generate_UnderCommaDecimalCulture_ProducesValidFile(bool ods) => await WithCommaDecimalCultureAsync(async () =>
    {
        using var spreadsheet = BuildWorkbook();
        var bytes = ods ? await spreadsheet.GenerateOdsFileAsync() : await spreadsheet.GenerateXlsxFileAsync();

        var result = ods ? SpreadsheetValidator.ValidateOds(bytes) : SpreadsheetValidator.ValidateXlsx(bytes);
        Assert.True(result.Issues.Count == 0, result.ToString());

        using var importer = new Spreadsheet();
        var sheet = (ods ? await importer.ImportOdsFileAsync(bytes) : await importer.ImportXlsxFileAsync(bytes)).Sheets[0];
        Assert.Equal("2024-01-31T12:30:00", sheet.GetCell(1, 1)?.Value);
        Assert.Equal(14.5, sheet.GetCell(1, 2)?.Style?.FontSize);
        Assert.Equal(20.5, sheet.ColumnWidths[1], 1);
    });

    /// <summary>
    /// A file generated under the invariant culture imports identically under a comma-decimal culture.
    /// </summary>
    [Fact]
    public async Task Import_UnderCommaDecimalCulture_ReadsDatesAndNumbers()
    {
        using var spreadsheet = BuildWorkbook();
        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        await WithCommaDecimalCultureAsync(async () =>
        {
            using var importer = new Spreadsheet();
            var sheet = (await importer.ImportXlsxFileAsync(bytes)).Sheets[0];

            Assert.Equal("2024-01-31T12:30:00", sheet.GetCell(1, 1)?.Value);
            Assert.Equal(CellValueType.DateTime, sheet.GetCell(1, 1)?.ValueType);
            Assert.Equal("1234.5", sheet.GetCell(1, 3)?.Value);
        });
    }
}
