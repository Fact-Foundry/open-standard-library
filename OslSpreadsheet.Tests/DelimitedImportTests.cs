using System.Text;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for importing and streaming delimited files: RFC 4180 quoting, each delimiter, and file encodings.
/// </summary>
public class DelimitedImportTests
{
    /// <summary>
    /// Imports text with the given delimiter and encoding and returns the first sheet's rows as arrays.
    /// </summary>
    private static async Task<List<string[]>> ImportAsync(byte[] bytes, ColumnDelimeter delimiter = ColumnDelimeter.Comma, FileEncoding encoding = FileEncoding.UTF8)
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.ColumnDelimeter = delimiter;
        spreadsheet.Workbook.FileEncoding = encoding;

        var sheet = (await spreadsheet.ImportCsvFileAsync(bytes)).Sheets[0];

        return sheet.Cells.GroupBy(c => c.Row).OrderBy(g => g.Key)
            .Select(g => g.OrderBy(c => c.Column).Select(c => c.Value).ToArray())
            .ToList();
    }

    private static Task<List<string[]>> ImportAsync(string text, ColumnDelimeter delimiter = ColumnDelimeter.Comma) =>
        ImportAsync(Encoding.UTF8.GetBytes(text), delimiter);

    /// <summary>
    /// Unquoted CSV imports correctly.
    /// </summary>
    [Fact]
    public async Task Csv_Unquoted_Imports()
    {
        var rows = await ImportAsync("Name,Score\nAlice,95\n");

        Assert.Equal(new[] { "Name", "Score" }, rows[0]);
        Assert.Equal(new[] { "Alice", "95" }, rows[1]);
    }

    /// <summary>
    /// Quoted values may contain delimiters, doubled quotes, and line breaks; quoted and unquoted values can be mixed.
    /// </summary>
    [Fact]
    public async Task Csv_Rfc4180Quoting_Imports()
    {
        var rows = await ImportAsync("Name,Quote,Notes\r\n\"Smith, John\",\"He said \"\"hi\"\"\",\"line one\r\nline two\"\r\nJane,plain,\r\n");

        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "Smith, John", "He said \"hi\"", "line one\r\nline two" }, rows[1]);
        Assert.Equal(new[] { "Jane", "plain", "" }, rows[2]);
    }

    /// <summary>
    /// Blank lines are skipped and a trailing line break does not create an extra row.
    /// </summary>
    [Fact]
    public async Task Csv_BlankLines_AreSkipped()
    {
        var rows = await ImportAsync("a,b\n\n1,2\n\n");

        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { "1", "2" }, rows[1]);
    }

    /// <summary>
    /// Malformed quoting is imported leniently rather than throwing or dropping data.
    /// </summary>
    [Fact]
    public async Task Csv_MalformedQuoting_IsLenient()
    {
        var rows = await ImportAsync("Item,Size\nTV,55\" screen\n");

        Assert.Equal(new[] { "TV", "55\" screen" }, rows[1]);
    }

    /// <summary>
    /// The workbook's ColumnDelimeter selects the delimiter used on import.
    /// </summary>
    [Theory]
    [InlineData(ColumnDelimeter.Tab, "Name\tScore\nAlice\t95\n")]
    [InlineData(ColumnDelimeter.Pipe, "Name|Score\nAlice|95\n")]
    [InlineData(ColumnDelimeter.ASCII, "Name\u001FScore\u001EAlice\u001F95\u001E")]
    public async Task Import_UsesWorkbookDelimiter(ColumnDelimeter delimiter, string text)
    {
        var rows = await ImportAsync(text, delimiter);

        Assert.Equal(new[] { "Name", "Score" }, rows[0]);
        Assert.Equal(new[] { "Alice", "95" }, rows[1]);
    }

    /// <summary>
    /// Values containing the delimiter, line breaks, or quotes survive a round-trip for every delimiter.
    /// </summary>
    [Theory]
    [InlineData(ColumnDelimeter.Comma)]
    [InlineData(ColumnDelimeter.Tab)]
    [InlineData(ColumnDelimeter.Pipe)]
    public async Task RoundTrip_PreservesAwkwardValues(ColumnDelimeter delimiter)
    {
        var values = new[] { "a,b", "tab\there", "pipe|here", "line\nbreak", "\"quoted\"", "5\" Fitting", "" , "end" };

        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.ColumnDelimeter = delimiter;
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        for (int c = 0; c < values.Length; c++)
            sheet.AddCell(1, c + 1, values[c]);

        var bytes = await spreadsheet.GenerateCsvFileAsync();
        var rows = await ImportAsync(bytes, delimiter);

        Assert.Equal(values, Assert.Single(rows));
    }

    /// <summary>
    /// Tab and pipe files only quote values that would otherwise be misread.
    /// </summary>
    [Fact]
    public async Task Generate_Tab_QuotesOnlyWhenNeeded()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.ColumnDelimeter = ColumnDelimeter.Tab;
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "5\" Fitting");
        sheet.AddCell(1, 2, "a\tb");

        var text = Encoding.UTF8.GetString(await spreadsheet.GenerateCsvFileAsync());

        Assert.Equal("5\" Fitting\t\"a\tb\"" + Environment.NewLine, text);
    }

    /// <summary>
    /// The workbook's FileEncoding is used on import.
    /// </summary>
    [Theory]
    [InlineData(FileEncoding.Unicode)]
    [InlineData(FileEncoding.UTF32)]
    public async Task Import_UsesWorkbookEncoding(FileEncoding fileEncoding)
    {
        Encoding encoding = fileEncoding == FileEncoding.Unicode ? new UnicodeEncoding(false, false) : new UTF32Encoding(false, false);
        var bytes = encoding.GetBytes("Name,City\nRenée,Zürich\n");

        var rows = await ImportAsync(bytes, ColumnDelimeter.Comma, fileEncoding);

        Assert.Equal(new[] { "Renée", "Zürich" }, rows[1]);
    }

    /// <summary>
    /// A UTF-8 byte order mark is not included in the first value.
    /// </summary>
    [Fact]
    public async Task Import_StripsByteOrderMark()
    {
        var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes("Name,Score\n")).ToArray();

        var rows = await ImportAsync(bytes);

        Assert.Equal("Name", rows[0][0]);
    }

    /// <summary>
    /// The imported workbook keeps the delimiter and encoding it was imported with.
    /// </summary>
    [Fact]
    public async Task Import_KeepsSettingsOnWorkbook()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.ColumnDelimeter = ColumnDelimeter.Pipe;
        spreadsheet.Workbook.FileEncoding = FileEncoding.ASCII;

        var workbook = await spreadsheet.ImportCsvFileAsync(Encoding.ASCII.GetBytes("a|b\n"));

        Assert.Equal(ColumnDelimeter.Pipe, workbook.ColumnDelimeter);
        Assert.Equal(FileEncoding.ASCII, workbook.FileEncoding);
    }

    /// <summary>
    /// The streaming reader handles quoted values spanning lines and uses the workbook's delimiter.
    /// </summary>
    [Fact]
    public async Task ReadCsvRows_HandlesMultilineValuesAndDelimiter()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.ColumnDelimeter = ColumnDelimeter.Tab;
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Name\tNotes\nAlice\t\"line one\nline two\"\nBob\tok\n"));

        var rows = new List<string[]>();
        await foreach (var row in spreadsheet.ReadCsvRowsAsync(stream, hasHeaderRow: true))
            rows.Add(row);

        Assert.Equal(new[] { "Name", "Notes" }, spreadsheet.CsvHeaders);
        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { "Alice", "line one\nline two" }, rows[0]);
        Assert.Equal(new[] { "Bob", "ok" }, rows[1]);
    }
}
