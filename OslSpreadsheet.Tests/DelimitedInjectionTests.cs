using System.Text;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for protection against CSV injection: text values that a spreadsheet application would evaluate as formulas
/// when opening a delimited file.
/// </summary>
public class DelimitedInjectionTests
{
    private static async Task<string> GenerateCsvAsync(bool escape, params (string Value, CellValueType Type)[] cells)
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.EscapeFormulaLikeValues = escape;
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        for (int i = 0; i < cells.Length; i++)
            sheet.AddCell(1, i + 1, cells[i].Value, cells[i].Type);
        return Encoding.UTF8.GetString(await spreadsheet.GenerateCsvFileAsync()).TrimEnd();
    }

    /// <summary>
    /// With escaping on, formula-like text values are prefixed with an apostrophe.
    /// </summary>
    [Theory]
    [InlineData("=cmd|'/c calc'!A0")]
    [InlineData("+1+1")]
    [InlineData("-1+cmd|'/c calc'!A0")]
    [InlineData("@SUM(1)")]
    [InlineData("\t=1+1")]
    [InlineData("\r=1+1")]
    public async Task Escape_PrefixesFormulaLikeText(string value)
    {
        var csv = await GenerateCsvAsync(escape: true, (value, CellValueType.String));

        Assert.Equal($"\"'{value}\"", csv);
    }

    /// <summary>
    /// Plain numbers starting with a sign, numeric cells, and ordinary text are left alone.
    /// </summary>
    [Fact]
    public async Task Escape_LeavesNumbersAndOrdinaryTextAlone()
    {
        var csv = await GenerateCsvAsync(escape: true,
            ("-5", CellValueType.String), ("+3.5", CellValueType.String), ("-1e3", CellValueType.String),
            ("-42", CellValueType.Float), ("hello", CellValueType.String), ("a=b", CellValueType.String));

        Assert.Equal("\"-5\",\"+3.5\",\"-1e3\",\"-42\",\"hello\",\"a=b\"", csv);
    }

    /// <summary>
    /// Escaping is off by default, so values are written verbatim.
    /// </summary>
    [Fact]
    public async Task Escape_OffByDefault_WritesVerbatim()
    {
        var csv = await GenerateCsvAsync(escape: false, ("=1+1", CellValueType.String));

        Assert.Equal("\"=1+1\"", csv);
    }

    /// <summary>
    /// Escaping applies to every delimiter.
    /// </summary>
    [Theory]
    [InlineData(ColumnDelimeter.Tab, "'=1+1")]
    [InlineData(ColumnDelimeter.Pipe, "'=1+1")]
    [InlineData(ColumnDelimeter.ASCII, "'=1+1")]
    public async Task Escape_AppliesToAllDelimiters(ColumnDelimeter delimiter, string expected)
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.EscapeFormulaLikeValues = true;
        spreadsheet.Workbook.ColumnDelimeter = delimiter;
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "=1+1");

        var text = Encoding.UTF8.GetString(await spreadsheet.GenerateCsvFileAsync()).TrimEnd('\r', '\n', '\u001E');

        Assert.Equal(expected, text);
    }

    /// <summary>
    /// Validate() warns about formula-like text when escaping is off, and is quiet when it is on or the value is a number.
    /// </summary>
    [Fact]
    public void WorkbookValidate_WarnsUnlessEscaping()
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet("Data");
        sheet.AddCell(2, 1, "=HYPERLINK(\"http://evil\")");
        sheet.AddCell(2, 2, "-5");
        sheet.AddCell(2, 3, "-7", CellValueType.Float);

        var result = workbook.Validate(ValidationFileFormat.Delimited);
        var warning = Assert.Single(result.Warnings);
        Assert.Equal("WB_DELIMITED_FORMULA_LIKE_VALUE", warning.Code);
        Assert.Equal("A2", warning.Cell);
        Assert.True(result.IsValid);

        workbook.EscapeFormulaLikeValues = true;
        Assert.Empty(workbook.Validate(ValidationFileFormat.Delimited).Issues);

        // Not a concern for XLSX or ODS, where text cells are stored as text
        workbook.EscapeFormulaLikeValues = false;
        Assert.Empty(workbook.Validate(ValidationFileFormat.Xlsx).Issues);
    }

    /// <summary>
    /// The file validator warns about formula-like fields in a delimited file, quoted or not, but not about negative numbers.
    /// </summary>
    [Fact]
    public void FileValidator_WarnsAboutFormulaLikeFields()
    {
        var csv = "Name,Amount,Note\nAlice,-5,\"=cmd|'/c calc'!A0\"\nBob,+12.50,@SUM(A1)\nCarol,-1234567890123456789,ok\n"u8.ToArray();

        var result = SpreadsheetValidator.ValidateDelimited(csv);

        Assert.True(result.IsValid);
        var warnings = result.Warnings.Where(w => w.Code == "CSV_FORMULA_LIKE_VALUE").ToList();
        Assert.Equal(2, warnings.Count);
        Assert.Equal(2, warnings[0].Row);
        Assert.Equal(3, warnings[1].Row);
    }

    /// <summary>
    /// An escaped value imports back with its apostrophe; this library doesn't strip it the way Excel hides it.
    /// </summary>
    [Fact]
    public async Task Escaped_RoundTrip_KeepsApostrophe()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.EscapeFormulaLikeValues = true;
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "=1+1");
        var bytes = await spreadsheet.GenerateCsvFileAsync();

        using var importer = new Spreadsheet();
        var cell = (await importer.ImportCsvFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Equal("'=1+1", cell.Value);
    }
}
