using System.Globalization;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for <see cref="oCell.GetFormattedValue"/>, which applies <see cref="CellStyle.NumberFormat"/> to a cell's value.
/// </summary>
public class FormattedValueTests
{
    private static oCell Cell(string value, CellValueType type, string? format = null) =>
        new(1, 1) { Value = value, ValueType = type, Style = format == null ? null : new CellStyle { NumberFormat = format } };

    // --- Numbers ---

    [Theory]
    [InlineData("#,##0.00", "1234.5", "1,234.50")]
    [InlineData("0", "1234.5", "1235")]
    [InlineData("0", "2.5", "3")]
    [InlineData("0.00", "2.675", "2.68")]
    [InlineData("0.00", "0.125", "0.13")]
    [InlineData("#,##0", "1234567", "1,234,567")]
    [InlineData("#,###", "0", "")]
    [InlineData("0.00", "-1234.5", "-1234.50")]
    [InlineData("0.00", "-0.001", "0.00")]
    [InlineData("#.00", "0.5", ".50")]
    [InlineData("000", "7", "007")]
    [InlineData("0.0#", "1.5", "1.5")]
    [InlineData("0.0#", "1.567", "1.57")]
    [InlineData("0.##", "1", "1.")]
    [InlineData("0", "1E+20", "100000000000000000000")]
    public void Number(string format, string value, string expected)
    {
        Assert.Equal(expected, Cell(value, CellValueType.Float, format).GetFormattedValue());
    }

    [Fact]
    public void Number_Int64Cell()
    {
        Assert.Equal("1,234,567", Cell("1234567", CellValueType.Int64, "#,##0").GetFormattedValue());
    }

    [Theory]
    [InlineData("0%", "0.256", "26%")]
    [InlineData("0.0%", "0.256", "25.6%")]
    [InlineData("0.00%", "-0.05", "-5.00%")]
    [InlineData("0.0 %", "0.5", "50.0 %")]
    public void Percentage(string format, string value, string expected)
    {
        Assert.Equal(expected, Cell(value, CellValueType.Float, format).GetFormattedValue());
    }

    [Theory]
    [InlineData("\"$\"#,##0.00", "1234.5", "$1,234.50")]
    [InlineData("\"$\"#,##0.00", "-9.99", "-$9.99")]
    [InlineData("#,##0.00 \"EUR\"", "1234.5", "1,234.50 EUR")]
    [InlineData("[$€-407]#,##0.00", "1234.5", "€1,234.50")]
    [InlineData("#,##0.00 [$€-407]", "1234.5", "1,234.50 €")]
    [InlineData("\\$0.00", "5", "$5.00")]
    [InlineData("#,##0.00_)", "5", "5.00")]
    [InlineData("_(0.00_)", "5", "5.00")]
    [InlineData("0.00_) \"x\"", "5", "5.00  x")]
    [InlineData("\"Total: \"0", "42", "Total: 42")]
    public void CurrencyAndLiterals(string format, string value, string expected)
    {
        Assert.Equal(expected, Cell(value, CellValueType.Float, format).GetFormattedValue());
    }

    // --- Sections ---

    private const string ExcelCurrency = "\"$\"#,##0.00_);(\"$\"#,##0.00)";
    private const string ExcelCurrencyRed = "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
    private const string ExcelAccounting = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)";

    [Theory]
    [InlineData(ExcelCurrency, "1234.5", "$1,234.50")]
    [InlineData(ExcelCurrency, "-1234.5", "($1,234.50)")]
    [InlineData(ExcelCurrency, "0", "$0.00")]
    [InlineData(ExcelCurrencyRed, "-1234.5", "($1,234.50)")]
    [InlineData(ExcelAccounting, "1234.5", "$1,234.50")]
    [InlineData(ExcelAccounting, "-1234.5", "$(1,234.50)")]
    [InlineData(ExcelAccounting, "0", "$-")]
    [InlineData("#,##0.00;(#,##0.00)", "-1234.5", "(1,234.50)")]
    [InlineData("#,##0.00;-#,##0.00", "-1234.5", "-1,234.50")]
    [InlineData("0;0", "-5", "5")]
    [InlineData("[Red]0.00", "5", "5.00")]
    [InlineData("0.00;[Red]-0.00", "-5", "-5.00")]
    [InlineData("0;-0;\"zero\"", "0", "zero")]
    [InlineData("0;-0;\"zero\"", "3", "3")]
    [InlineData("0;-0;\"zero\";\"text: \"@", "-3", "-3")]
    public void Sections_ChosenBySign(string format, string value, string expected)
    {
        Assert.Equal(expected, Cell(value, CellValueType.Float, format).GetFormattedValue());
    }

    /// <summary>
    /// Excel's built-in formats 37-40 all have a negative section.
    /// </summary>
    [Theory]
    [InlineData(37, "-1234", "(1,234)")]
    [InlineData(38, "-1234", "(1,234)")]
    [InlineData(39, "-1234.5", "(1,234.50)")]
    [InlineData(40, "-1234.5", "(1,234.50)")]
    [InlineData(37, "1234", "1,234 ")] // 37 is "#,##0 ;(#,##0)": a literal space, which is kept, unlike _) padding
    [InlineData(39, "1234.5", "1,234.50")]
    public void Sections_BuiltInFormats(int id, string value, string expected)
    {
        var code = OslSpreadsheet.Services.NumberFormatTranslator.BuiltInFormats[id];

        Assert.Equal(expected, Cell(value, CellValueType.Float, code).GetFormattedValue());
    }

    [Theory]
    [InlineData("m/d/yy;@", "3/15/21")]
    [InlineData("[$-409]mmmm d, yyyy;@", "March 15, 2021")]
    [InlineData("[$-F800]dddd, mmmm dd, yyyy", "Monday, March 15, 2021")]
    public void Sections_DateWithTextSection(string format, string expected)
    {
        Assert.Equal(expected, Cell("2021-03-15T00:00:00", CellValueType.DateTime, format).GetFormattedValue());
    }

    [Theory]
    [InlineData("\"ID-\"@", "ID-00123")]
    [InlineData("@\" units\"", "00123 units")]
    [InlineData("m/d/yy;@", "00123")]
    [InlineData(ExcelAccounting, "00123")]
    [InlineData("0;-0;0;\"<\"@\">\"", "<00123>")]
    [InlineData("0.00", "00123")]
    [InlineData("0;-0;0;\"fixed\"", "fixed")]
    public void Sections_TextCell(string format, string expected)
    {
        Assert.Equal(expected, Cell("00123", CellValueType.String, format).GetFormattedValue());
    }

    [Fact]
    public void Sections_NumberWithOnlyTextSectionUsesDefault()
    {
        Assert.Equal("1234.5", Cell("1234.5", CellValueType.Float, "\"ID-\"@").GetFormattedValue());
        Assert.Equal("2024-01-31", Cell("2024-01-31", CellValueType.DateTime, "@").GetFormattedValue());
    }

    // --- Dates and times ---

    [Theory]
    [InlineData("mm/dd/yyyy", "2021-03-15T00:00:00", "03/15/2021")]
    [InlineData("m/d/yy", "2021-03-05T00:00:00", "3/5/21")]
    [InlineData("d-mmm-yy", "2024-01-31", "31-Jan-24")]
    [InlineData("mmmm d, yyyy", "2024-01-31", "January 31, 2024")]
    [InlineData("mmmmm", "2024-01-31", "J")]
    [InlineData("ddd", "2024-01-31", "Wed")]
    [InlineData("dddd", "2024-01-31", "Wednesday")]
    [InlineData("yyyy-mm-dd hh:mm", "2024-01-31T13:45:00", "2024-01-31 13:45")]
    [InlineData("h:mm AM/PM", "2024-01-31T13:45:00", "1:45 PM")]
    [InlineData("h:mm AM/PM", "2024-01-31T00:05:00", "12:05 AM")]
    [InlineData("hh:mm:ss", "2024-01-31T09:05:03", "09:05:03")]
    [InlineData("h:m:s", "2024-01-31T09:05:03", "9:5:3")]
    [InlineData("mm:ss", "2024-01-31T09:05:03", "05:03")]
    [InlineData("yyyy\"年\"m\"月\"", "2024-01-31", "2024年1月")]
    public void Date(string format, string value, string expected)
    {
        Assert.Equal(expected, Cell(value, CellValueType.DateTime, format).GetFormattedValue());
    }

    [Fact]
    public void Date_FloatSerialWithDateFormat()
    {
        Assert.Equal("01/31/2024", Cell("45322", CellValueType.Float, "mm/dd/yyyy").GetFormattedValue());
    }

    [Fact]
    public void Date_WithNumberFormatShowsSerial()
    {
        Assert.Equal("45322.50", Cell("2024-01-31T12:00:00", CellValueType.DateTime, "0.00").GetFormattedValue());
    }

    // --- Defaults ---

    [Theory]
    [InlineData("2024-01-31T00:00:00", "2024-01-31")]
    [InlineData("2024-01-31", "2024-01-31")]
    [InlineData("2024-01-31T13:45:00", "2024-01-31 13:45:00")]
    public void Default_DateTime(string value, string expected)
    {
        Assert.Equal(expected, Cell(value, CellValueType.DateTime).GetFormattedValue());
        Assert.Equal(expected, Cell(value, CellValueType.DateTime, "General").GetFormattedValue());
    }

    [Theory]
    [InlineData("true", "TRUE")]
    [InlineData("False", "FALSE")]
    [InlineData("1", "TRUE")]
    [InlineData("0", "FALSE")]
    public void Boolean_IgnoresFormat(string value, string expected)
    {
        Assert.Equal(expected, Cell(value, CellValueType.Boolean).GetFormattedValue());
        Assert.Equal(expected, Cell(value, CellValueType.Boolean, "0.00").GetFormattedValue());
    }

    [Theory]
    [InlineData("1234.5", CellValueType.Float)]
    [InlineData("hello", CellValueType.String)]
    [InlineData("", CellValueType.Float)]
    public void Default_ReturnsValue(string value, CellValueType type)
    {
        Assert.Equal(value, Cell(value, type).GetFormattedValue());
    }

    [Fact]
    public void Text_FormatReturnsValue()
    {
        Assert.Equal("00123", Cell("00123", CellValueType.String, "@").GetFormattedValue());
    }

    [Fact]
    public void String_WithNumberFormatReturnsValue()
    {
        Assert.Equal("1234.5", Cell("1234.5", CellValueType.String, "#,##0.00").GetFormattedValue());
    }

    // --- Fallbacks ---

    [Theory]
    [InlineData("[>100]0")]
    [InlineData("[<0]\"neg\";0")]
    [InlineData("0.00E+00")]
    [InlineData("# ?/?")]
    [InlineData("[h]:mm:ss")]
    [InlineData("#,##0,")]
    [InlineData("ss.0")]
    public void Unsupported_ReturnsValue(string format)
    {
        Assert.Equal("1234.5", Cell("1234.5", CellValueType.Float, format).GetFormattedValue());
    }

    [Theory]
    [InlineData("abc", CellValueType.Float, "0.00")]
    [InlineData("not a date", CellValueType.DateTime, "mm/dd/yyyy")]
    [InlineData("1e300", CellValueType.Float, "0")]
    [InlineData("1e300", CellValueType.Float, "mm/dd/yyyy")]
    public void MismatchedValue_ReturnsValue(string value, CellValueType type, string format)
    {
        Assert.Equal(value, Cell(value, type, format).GetFormattedValue());
    }

    // --- Culture ---

    [Fact]
    public void Culture_DefaultsToInvariant()
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");

            Assert.Equal("1,234.50", Cell("1234.5", CellValueType.Float, "#,##0.00").GetFormattedValue());
            Assert.Equal("31 Jan 2024", Cell("2024-01-31", CellValueType.DateTime, "d mmm yyyy").GetFormattedValue());
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void Culture_SuppliesSeparatorsAndNames()
    {
        var german = new CultureInfo("de-DE");

        Assert.Equal("1.234,50", Cell("1234.5", CellValueType.Float, "#,##0.00").GetFormattedValue(german));
        Assert.Equal("Mittwoch, 31. Januar 2024", Cell("2024-01-31", CellValueType.DateTime, "dddd, d. mmmm yyyy").GetFormattedValue(german));
    }

    // --- Imported files ---

    [Fact]
    public async Task ImportedXlsx_FormatsWithImportedFormat()
    {
        using var source = new Spreadsheet();
        var sheet = source.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "1234.5", CellValueType.Float).Style = new CellStyle { NumberFormat = "\"$\"#,##0.00" };
        sheet.AddCell(1, 2, "2021-03-15", CellValueType.DateTime).Style = new CellStyle { NumberFormat = "mm/dd/yyyy" };
        sheet.AddCell(1, 3, "2021-03-15", CellValueType.DateTime);
        var bytes = await source.GenerateXlsxFileAsync();

        using var spreadsheet = new Spreadsheet();
        var cells = (await spreadsheet.ImportXlsxFileAsync(bytes)).Sheets[0];

        Assert.Equal("$1,234.50", cells.GetCell(1, 1)!.GetFormattedValue());
        Assert.Equal("03/15/2021", cells.GetCell(1, 2)!.GetFormattedValue());
        Assert.Equal("2021-03-15", cells.GetCell(1, 3)!.GetFormattedValue());
    }

    [Fact]
    public async Task ImportedOds_FormatsWithImportedFormat()
    {
        using var source = new Spreadsheet();
        var sheet = source.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "0.256", CellValueType.Float).Style = new CellStyle { NumberFormat = "0.0%" };
        sheet.AddCell(1, 2, "2024-01-31T13:45:00", CellValueType.DateTime).Style = new CellStyle { NumberFormat = "h:mm AM/PM" };
        var bytes = await source.GenerateOdsFileAsync();

        using var spreadsheet = new Spreadsheet();
        var cells = (await spreadsheet.ImportOdsFileAsync(bytes)).Sheets[0];

        Assert.Equal("25.6%", cells.GetCell(1, 1)!.GetFormattedValue());
        Assert.Equal("1:45 PM", cells.GetCell(1, 2)!.GetFormattedValue());
    }
}
