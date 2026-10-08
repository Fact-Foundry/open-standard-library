using System.IO.Compression;
using System.Xml.Linq;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Services;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for CellStyle.NumberFormat: translation between Excel format codes and ODS data styles, and round-trips through both file formats.
/// </summary>
public class NumberFormatTests
{
    private static string ReadEntry(byte[] package, string path)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry(path)!.Open());
        return reader.ReadToEnd();
    }

    private static Spreadsheet BuildFormatted(string format, string value, CellValueType type)
    {
        var spreadsheet = new Spreadsheet();
        var cell = spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, value, type);
        cell.Style = new CellStyle { NumberFormat = format };
        return spreadsheet;
    }

    // --- Translator ---

    /// <summary>
    /// Excel format codes translate to the expected ODS data style elements.
    /// </summary>
    [Theory]
    [InlineData("#,##0.00", "<number:number-style style:name=\"N\"><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/></number:number-style>")]
    [InlineData("0", "<number:number-style style:name=\"N\"><number:number number:decimal-places=\"0\" number:min-decimal-places=\"0\" number:min-integer-digits=\"1\"/></number:number-style>")]
    [InlineData("0.0%", "<number:percentage-style style:name=\"N\"><number:number number:decimal-places=\"1\" number:min-decimal-places=\"1\" number:min-integer-digits=\"1\"/><number:text>%</number:text></number:percentage-style>")]
    [InlineData("\"$\"#,##0.00", "<number:currency-style style:name=\"N\"><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/></number:currency-style>")]
    [InlineData("#,##0 \"EUR\"", "<number:currency-style style:name=\"N\"><number:number number:decimal-places=\"0\" number:min-decimal-places=\"0\" number:min-integer-digits=\"1\" number:grouping=\"true\"/><number:currency-symbol> EUR</number:currency-symbol></number:currency-style>")]
    [InlineData("mm/dd/yyyy", "<number:date-style style:name=\"N\"><number:month number:style=\"long\"/><number:text>/</number:text><number:day number:style=\"long\"/><number:text>/</number:text><number:year number:style=\"long\"/></number:date-style>")]
    [InlineData("mmm d, yyyy", "<number:date-style style:name=\"N\"><number:month number:style=\"short\" number:textual=\"true\"/><number:text> </number:text><number:day number:style=\"short\"/><number:text>, </number:text><number:year number:style=\"long\"/></number:date-style>")]
    [InlineData("h:mm AM/PM", "<number:time-style style:name=\"N\"><number:hours number:style=\"short\"/><number:text>:</number:text><number:minutes number:style=\"long\"/><number:text> </number:text><number:am-pm/></number:time-style>")]
    [InlineData("@", "<number:text-style style:name=\"N\"><number:text-content/></number:text-style>")]
    public void ToOdsDataStyle_TranslatesExcelCodes(string code, string expected)
    {
        Assert.Equal(expected, NumberFormatTranslator.ToOdsDataStyle(code, "N"));
    }

    /// <summary>
    /// Codes with no ODS equivalent return null rather than a wrong translation.
    /// </summary>
    [Theory]
    [InlineData("0.00E+00")]
    [InlineData("[h]:mm:ss")]
    [InlineData("#,##0,")]
    [InlineData("# ?/?")]
    public void ToOdsDataStyle_UnsupportedCodes_ReturnNull(string code)
    {
        Assert.Null(NumberFormatTranslator.ToOdsDataStyle(code, "N"));
    }

    /// <summary>
    /// An 'm' between hours and seconds is minutes; elsewhere it is the month.
    /// </summary>
    [Fact]
    public void Parse_DistinguishesMinutesFromMonths()
    {
        var ods = NumberFormatTranslator.ToOdsDataStyle("yyyy-mm-dd hh:mm:ss", "N")!;

        Assert.Contains("<number:month number:style=\"long\"/>", ods);
        Assert.Contains("<number:minutes number:style=\"long\"/>", ods);
    }

    /// <summary>
    /// ODS data styles written by other applications translate back to Excel codes, and the General style maps to none.
    /// </summary>
    [Theory]
    [InlineData("<number:number-style style:name=\"N\"><number:number number:decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/></number:number-style>", "#,##0.00")]
    [InlineData("<number:percentage-style style:name=\"N\"><number:number number:decimal-places=\"0\" number:min-integer-digits=\"1\"/><number:text>%</number:text></number:percentage-style>", "0%")]
    [InlineData("<number:currency-style style:name=\"N\"><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/></number:currency-style>", "\"$\"#,##0.00")]
    [InlineData("<number:date-style style:name=\"N\"><number:day/><number:text>.</number:text><number:month/><number:text>.</number:text><number:year number:style=\"long\"/></number:date-style>", "d.m.yyyy")]
    [InlineData("<number:date-style style:name=\"N\"><number:day-of-week number:style=\"long\"/><number:text>, </number:text><number:month number:style=\"long\" number:textual=\"true\"/><number:text> </number:text><number:day/></number:date-style>", "dddd, mmmm d")]
    [InlineData("<number:text-style style:name=\"N\"><number:text-content/></number:text-style>", "@")]
    [InlineData("<number:number-style style:name=\"N0\"><number:number number:min-integer-digits=\"1\"/></number:number-style>", null)]
    public void FromOdsDataStyle_TranslatesOdsStyles(string xml, string? expected)
    {
        var wrapped = "<r xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\">" + xml + "</r>";
        var element = XElement.Parse(wrapped, LoadOptions.PreserveWhitespace).Elements().First();
        Assert.Equal(expected, NumberFormatTranslator.FromOdsDataStyle(element));
    }

    // --- XLSX ---

    /// <summary>
    /// A custom format is written to styles.xml and the cell references it; built-in codes use Excel's fixed ids.
    /// </summary>
    [Fact]
    public async Task Xlsx_WritesNumFmtAndCellStyle()
    {
        using var spreadsheet = BuildFormatted("\"$\"#,##0.00", "9.99", CellValueType.Float);
        spreadsheet.Workbook.Sheets[0].AddCell(1, 2, "0.25", CellValueType.Float).Style = new CellStyle { NumberFormat = "0%" };

        var bytes = await spreadsheet.GenerateXlsxFileAsync();
        var styles = ReadEntry(bytes, "xl/styles.xml");
        var sheet = ReadEntry(bytes, "xl/worksheets/sheet1.xml");

        Assert.Contains("<numFmt numFmtId=\"166\" formatCode=\"&quot;$&quot;#,##0.00\"/>", styles);
        Assert.Contains("numFmtId=\"166\"", styles);
        Assert.Contains("numFmtId=\"9\"", styles); // built-in 0%
        Assert.DoesNotContain("formatCode=\"0%\"", styles);
        Assert.Matches("<c r=\"A1\" s=\"[1-9]\"><v>9.99</v></c>", sheet);
        Assert.True(SpreadsheetValidator.ValidateXlsx(bytes).Issues.Count == 0);
    }

    /// <summary>
    /// Every supported format code survives an XLSX round-trip unchanged.
    /// </summary>
    [Theory]
    [InlineData("#,##0.00", "1234.5", CellValueType.Float)]
    [InlineData("0.0%", "0.256", CellValueType.Float)]
    [InlineData("\"$\"#,##0.00", "9.99", CellValueType.Float)]
    [InlineData("[$€-407]#,##0.00", "1234.5", CellValueType.Float)]
    [InlineData("mmmm d, yyyy", "2024-01-31", CellValueType.DateTime)]
    [InlineData("h:mm AM/PM", "2024-01-31T13:45:00", CellValueType.DateTime)]
    [InlineData("0.00E+00", "12345", CellValueType.Float)]
    [InlineData("@", "007", CellValueType.String)]
    public async Task Xlsx_RoundTrip_PreservesNumberFormat(string format, string value, CellValueType type)
    {
        using var spreadsheet = BuildFormatted(format, value, type);
        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var cell = (await importer.ImportXlsxFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Equal(format, cell.Style?.NumberFormat);
        Assert.Equal(type, cell.ValueType);
    }

    /// <summary>
    /// Built-in formats in files from other applications (stored only by id) import with their code.
    /// </summary>
    [Fact]
    public async Task Xlsx_Import_ResolvesBuiltInFormatIds()
    {
        using var spreadsheet = BuildFormatted("0.00", "1.5", CellValueType.Float); // built-in id 2
        var bytes = await spreadsheet.GenerateXlsxFileAsync();
        Assert.DoesNotContain("formatCode", ReadEntry(bytes, "xl/styles.xml"));

        using var importer = new Spreadsheet();
        var cell = (await importer.ImportXlsxFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Equal("0.00", cell.Style?.NumberFormat);
    }

    /// <summary>
    /// DateTime cells without an explicit format use the defaults and come back with no Style, as before.
    /// </summary>
    [Fact]
    public async Task Xlsx_DefaultDateFormat_LeavesStyleNull()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "2024-01-31", CellValueType.DateTime);
        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var cell = (await importer.ImportXlsxFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Equal(CellValueType.DateTime, cell.ValueType);
        Assert.Null(cell.Style);
    }

    // --- ODS ---

    /// <summary>
    /// Custom formats are written as data styles, with percentage and currency cells given the matching ODS value type.
    /// </summary>
    [Fact]
    public async Task Ods_WritesDataStylesAndValueTypes()
    {
        using var spreadsheet = BuildFormatted("\"$\"#,##0.00", "9.99", CellValueType.Float);
        spreadsheet.Workbook.Sheets[0].AddCell(1, 2, "0.25", CellValueType.Float).Style = new CellStyle { NumberFormat = "0%" };

        var bytes = await spreadsheet.GenerateOdsFileAsync();
        var content = ReadEntry(bytes, "content.xml");

        Assert.Contains("<number:currency-style style:name=\"N100\">", content);
        Assert.Contains("<number:percentage-style style:name=\"N101\">", content);
        Assert.Contains("office:value-type=\"currency\" office:value=\"9.99\" office:currency=\"USD\"", content);
        Assert.Contains("office:value-type=\"percentage\" office:value=\"0.25\"", content);
        Assert.True(SpreadsheetValidator.ValidateOds(bytes).Issues.Count == 0, SpreadsheetValidator.ValidateOds(bytes).ToString());
    }

    /// <summary>
    /// Supported format codes survive an ODS round-trip.
    /// </summary>
    [Theory]
    [InlineData("#,##0.00", "1234.5", CellValueType.Float)]
    [InlineData("0.0%", "0.256", CellValueType.Float)]
    [InlineData("\"$\"#,##0.00", "9.99", CellValueType.Float)]
    [InlineData("mm/dd/yyyy", "2024-01-31", CellValueType.DateTime)]
    [InlineData("mmmm d, yyyy", "2024-01-31", CellValueType.DateTime)]
    [InlineData("h:mm AM/PM", "2024-01-31T13:45:00", CellValueType.DateTime)]
    [InlineData("@", "007", CellValueType.String)]
    public async Task Ods_RoundTrip_PreservesNumberFormat(string format, string value, CellValueType type)
    {
        using var spreadsheet = BuildFormatted(format, value, type);
        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var cell = (await importer.ImportOdsFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Equal(format, cell.Style?.NumberFormat);
        Assert.Equal(type, cell.ValueType);
        Assert.Equal(value, cell.Value);
    }

    /// <summary>
    /// A format ODS can't express is reported by Validate() as a warning and the cell falls back to the default style.
    /// </summary>
    [Fact]
    public async Task Ods_UnsupportedFormat_WarnsAndFallsBack()
    {
        using var spreadsheet = BuildFormatted("0.00E+00", "12345", CellValueType.Float);

        var result = spreadsheet.Workbook.Validate(ValidationFileFormat.Ods);
        Assert.True(result.IsValid);
        Assert.Equal("WB_NUMBER_FORMAT_UNSUPPORTED", Assert.Single(result.Warnings).Code);
        Assert.True(spreadsheet.Workbook.Validate(ValidationFileFormat.Xlsx).Issues.Count == 0);

        var bytes = await spreadsheet.GenerateOdsFileAsync();
        using var importer = new Spreadsheet();
        var cell = (await importer.ImportOdsFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Null(cell.Style);
        Assert.Equal("12345", cell.Value);
    }

    /// <summary>
    /// Cells with a visual style but no number format still import with no NumberFormat (the default N0 style is "General").
    /// </summary>
    [Fact]
    public async Task Ods_StyledCellWithoutFormat_ImportsNoNumberFormat()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "42", CellValueType.Float).Style = new CellStyle { Bold = true };
        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var cell = (await importer.ImportOdsFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Null(cell.Style?.NumberFormat);
    }

    /// <summary>
    /// A cell whose entire content is a single space is not dropped on import (whitespace-only XML text is preserved).
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Import_SingleSpaceCell_IsPreserved(bool ods)
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, " ");
        var bytes = ods ? await spreadsheet.GenerateOdsFileAsync() : await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = ods ? await importer.ImportOdsFileAsync(bytes) : await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(" ", workbook.Sheets[0].Cells.Single().Value);
    }
}
