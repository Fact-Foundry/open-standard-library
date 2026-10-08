using System.IO.Compression;
using System.Text;
using OoxSpreadsheet;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests that ODS import handles structures written by LibreOffice and other applications that the library itself doesn't generate.
/// </summary>
public class OdsImportTests
{
    /// <summary>
    /// Builds a minimal ODS package around the given table XML (the content of an office:spreadsheet element).
    /// </summary>
    private static byte[] BuildOds(string spreadsheetBody)
    {
        var content =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
            "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" " +
            "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\">" +
            SplitStyles(spreadsheetBody, out var body) +
            "<office:body><office:spreadsheet>" + body + "</office:spreadsheet></office:body>" +
            "</office:document-content>";

        var manifest =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\">" +
            "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>" +
            "<manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>" +
            "</manifest:manifest>";

        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            void Add(string name, string text, CompressionLevel level)
            {
                using var stream = archive.CreateEntry(name, level).Open();
                var bytes = Encoding.UTF8.GetBytes(text);
                stream.Write(bytes, 0, bytes.Length);
            }

            Add("mimetype", "application/vnd.oasis.opendocument.spreadsheet", CompressionLevel.NoCompression);
            Add("META-INF/manifest.xml", manifest, CompressionLevel.Fastest);
            Add("content.xml", content, CompressionLevel.Fastest);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Lets a test body start with an office:automatic-styles element, which belongs before office:body.
    /// </summary>
    private static string SplitStyles(string spreadsheetBody, out string body)
    {
        const string end = "</office:automatic-styles>";
        var idx = spreadsheetBody.IndexOf(end, StringComparison.Ordinal);
        if (!spreadsheetBody.StartsWith("<office:automatic-styles", StringComparison.Ordinal) || idx < 0)
        {
            body = spreadsheetBody;
            return "";
        }
        body = spreadsheetBody[(idx + end.Length)..];
        return spreadsheetBody[..(idx + end.Length)];
    }

    private static string StringCell(string value) =>
        $"<table:table-cell office:value-type=\"string\"><text:p>{value}</text:p></table:table-cell>";

    private static string Row(params string[] values) =>
        "<table:table-row>" + string.Concat(values.Select(StringCell)) + "</table:table-row>";

    /// <summary>
    /// Imports the package and returns the first sheet's values as (row, column, value) tuples in order.
    /// </summary>
    private static async Task<List<(int Row, int Col, string Value)>> ImportAsync(byte[] ods)
    {
        using var spreadsheet = new Spreadsheet();
        var workbook = await spreadsheet.ImportOdsFileAsync(ods);
        return workbook.Sheets[0].Cells.OrderBy(c => c.Row).ThenBy(c => c.Column).Select(c => (c.Row, c.Column, c.Value)).ToList();
    }

    /// <summary>
    /// Rows wrapped in table:table-header-rows (written by LibreOffice when rows are set to repeat on each printed page) are imported in order.
    /// </summary>
    [Fact]
    public async Task Import_ReadsRowsInsideHeaderRows()
    {
        var ods = BuildOds(
            "<table:table table:name=\"Data\">" +
            "<table:table-header-rows>" + Row("Name", "Score") + "</table:table-header-rows>" +
            Row("Alice", "95") +
            "</table:table>");

        var cells = await ImportAsync(ods);

        Assert.Equal(new[] { (1, 1, "Name"), (1, 2, "Score"), (2, 1, "Alice"), (2, 2, "95") }, cells);
    }

    /// <summary>
    /// Rows inside table:table-row-group (outline groups), including nested groups and table:table-rows wrappers, are imported in document order.
    /// </summary>
    [Fact]
    public async Task Import_ReadsRowsInsideNestedRowGroups()
    {
        var ods = BuildOds(
            "<table:table table:name=\"Data\">" +
            Row("r1") +
            "<table:table-row-group>" +
                Row("r2") +
                "<table:table-row-group>" + "<table:table-rows>" + Row("r3") + "</table:table-rows>" + "</table:table-row-group>" +
                Row("r4") +
            "</table:table-row-group>" +
            Row("r5") +
            "</table:table>");

        var cells = await ImportAsync(ods);

        Assert.Equal(new[] { (1, 1, "r1"), (2, 1, "r2"), (3, 1, "r3"), (4, 1, "r4"), (5, 1, "r5") }, cells);
    }

    /// <summary>
    /// Covered cells (hidden by a merge) occupy column positions, so cells after a merged range keep their columns.
    /// </summary>
    [Fact]
    public async Task Import_CoveredCells_KeepColumnPositions()
    {
        var ods = BuildOds(
            "<table:table table:name=\"Data\">" +
            "<table:table-row>" +
                "<table:table-cell office:value-type=\"string\" table:number-columns-spanned=\"3\" table:number-rows-spanned=\"1\"><text:p>Title</text:p></table:table-cell>" +
                "<table:covered-table-cell table:number-columns-repeated=\"2\"/>" +
                StringCell("after-merge") +
            "</table:table-row>" +
            Row("a", "b", "c", "d") +
            "</table:table>");

        var cells = await ImportAsync(ods);

        Assert.Equal(new[] { (1, 1, "Title"), (1, 4, "after-merge"), (2, 1, "a"), (2, 2, "b"), (2, 3, "c"), (2, 4, "d") }, cells);
    }

    /// <summary>
    /// Content inside a covered cell is hidden by the merge and is not imported.
    /// </summary>
    [Fact]
    public async Task Import_CoveredCellContent_IsIgnored()
    {
        var ods = BuildOds(
            "<table:table table:name=\"Data\">" +
            "<table:table-row>" +
                "<table:table-cell office:value-type=\"string\" table:number-columns-spanned=\"2\"><text:p>Merged</text:p></table:table-cell>" +
                "<table:covered-table-cell office:value-type=\"string\"><text:p>hidden</text:p></table:covered-table-cell>" +
                StringCell("visible") +
            "</table:table-row>" +
            "</table:table>");

        var cells = await ImportAsync(ods);

        Assert.Equal(new[] { (1, 1, "Merged"), (1, 3, "visible") }, cells);
    }

    /// <summary>
    /// Numeric cells import the stored office:value, not the formatted display text; percentage and currency cells are numeric too.
    /// </summary>
    [Fact]
    public async Task Import_NumericCells_UseStoredValueNotDisplayText()
    {
        var ods = BuildOds(
            "<table:table table:name=\"Data\"><table:table-row>" +
            "<table:table-cell office:value-type=\"float\" office:value=\"1234.5\"><text:p>1,234.50</text:p></table:table-cell>" +
            "<table:table-cell office:value-type=\"percentage\" office:value=\"0.25\"><text:p>25%</text:p></table:table-cell>" +
            "<table:table-cell office:value-type=\"currency\" office:currency=\"USD\" office:value=\"9.99\"><text:p>$9.99</text:p></table:table-cell>" +
            "</table:table-row></table:table>");

        using var spreadsheet = new Spreadsheet();
        var cells = (await spreadsheet.ImportOdsFileAsync(ods)).Sheets[0].Cells.OrderBy(c => c.Column).ToList();

        Assert.All(cells, c => Assert.Equal(OslSpreadsheet.Models.CellValueType.Float, c.ValueType));
        Assert.Equal(new[] { "1234.5", "0.25", "9.99" }, cells.Select(c => c.Value));
    }

    /// <summary>
    /// A cell with several paragraphs imports as multi-line text, and ODF whitespace elements expand to their characters.
    /// </summary>
    [Fact]
    public async Task Import_MultiParagraphAndWhitespaceElements_ArePreserved()
    {
        var ods = BuildOds(
            "<table:table table:name=\"Data\"><table:table-row>" +
            "<table:table-cell office:value-type=\"string\"><text:p>line one</text:p><text:p>line two</text:p></table:table-cell>" +
            "<table:table-cell office:value-type=\"string\"><text:p><text:s text:c=\"2\"/>padded<text:s/></text:p></table:table-cell>" +
            "<table:table-cell office:value-type=\"string\"><text:p>a<text:tab/>b<text:line-break/>c <text:span>bold</text:span></text:p></table:table-cell>" +
            "</table:table-row></table:table>");

        var cells = await ImportAsync(ods);

        Assert.Equal("line one\nline two", cells[0].Value);
        Assert.Equal("  padded ", cells[1].Value);
        Assert.Equal("a\tb\nc bold", cells[2].Value);
    }

    /// <summary>
    /// Cells without their own style take the column's default cell style, which LibreOffice uses when a whole column is formatted.
    /// Column widths are read, and the trailing filler column declaration is ignored.
    /// </summary>
    [Fact]
    public async Task Import_ColumnDefaultCellStyleAndWidths()
    {
        var ods = BuildOds(
            "<office:automatic-styles xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\">" +
            "<style:style style:name=\"co1\" style:family=\"table-column\"><style:table-column-properties style:column-width=\"5.08cm\"/></style:style>" +
            "<style:style style:name=\"co2\" style:family=\"table-column\"><style:table-column-properties style:column-width=\"2.258cm\"/></style:style>" +
            "<style:style style:name=\"ce1\" style:family=\"table-cell\"><style:table-cell-properties fo:background-color=\"#ffff00\" fo:border-bottom=\"1.76pt solid #1565c0\"/>" +
            "<style:text-properties fo:font-weight=\"bold\" fo:color=\"#ff0000\"/></style:style>" +
            "</office:automatic-styles>" +
            "<table:table table:name=\"Data\">" +
            "<table:table-column table:style-name=\"co1\" table:default-cell-style-name=\"ce1\"/>" +
            "<table:table-column table:style-name=\"co2\" table:number-columns-repeated=\"16383\"/>" +
            Row("Header", "plain") +
            "</table:table>");

        using var spreadsheet = new Spreadsheet();
        var sheet = (await spreadsheet.ImportOdsFileAsync(ods)).Sheets[0];

        var styled = sheet.GetCell(1, 1)?.Style;
        Assert.NotNull(styled);
        Assert.True(styled.Bold);
        Assert.Equal("#FF0000", styled.FontColor);
        Assert.Equal("#FFFF00", styled.BackgroundColor);
        Assert.Equal(OslSpreadsheet.Models.BorderStyle.Medium, styled.BorderBottom?.Style);
        Assert.Equal("#1565C0", styled.BorderBottom?.Color);
        Assert.Null(sheet.GetCell(1, 2)?.Style);

        Assert.Equal(25.29, sheet.ColumnWidths[1], 1);
        Assert.Single(sheet.ColumnWidths);
    }

    /// <summary>
    /// A repeated row inside a header-rows wrapper still advances the row index by its repeat count.
    /// </summary>
    [Fact]
    public async Task Import_RepeatedRowInsideHeaderRows_AdvancesRowIndex()
    {
        var ods = BuildOds(
            "<table:table table:name=\"Data\">" +
            "<table:table-header-rows><table:table-row table:number-rows-repeated=\"2\">" + StringCell("h") + "</table:table-row></table:table-header-rows>" +
            Row("data") +
            "</table:table>");

        var cells = await ImportAsync(ods);

        Assert.Equal(new[] { (1, 1, "h"), (2, 1, "h"), (3, 1, "data") }, cells);
    }
}
