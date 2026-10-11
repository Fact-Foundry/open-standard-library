using System.IO.Compression;
using System.Text;
using OoxSpreadsheet;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests that import rejects files designed to exhaust memory or crash the process: decompression bombs,
/// repeat-count expansion, DTD entity expansion, and deeply nested markup.
/// </summary>
public class ImportLimitTests
{
    private const string Namespaces =
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
        "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" " +
        "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"";

    /// <summary>
    /// Builds an ODS package whose content.xml is exactly the given text (so tests can include a DTD or odd structure).
    /// </summary>
    private static byte[] OdsWithContent(string contentXml)
    {
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
            Add("content.xml", contentXml, CompressionLevel.SmallestSize);
        }
        return ms.ToArray();
    }

    private static byte[] OdsWithRows(string rows) => OdsWithContent(
        $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content {Namespaces} office:version=\"1.3\">" +
        $"<office:body><office:spreadsheet><table:table table:name=\"T\">{rows}</table:table></office:spreadsheet></office:body></office:document-content>");

    /// <summary>
    /// A cell repeated across the whole sheet in a few bytes is stopped by MaxCells rather than expanded.
    /// </summary>
    [Fact]
    public async Task Ods_RepeatExpansion_StoppedByMaxCells()
    {
        var ods = OdsWithRows(
            "<table:table-row table:number-rows-repeated=\"1048576\">" +
            "<table:table-cell office:value-type=\"string\" table:number-columns-repeated=\"16384\"><text:p>x</text:p></table:table-cell>" +
            "</table:table-row>");
        Assert.True(ods.Length < 2000);

        using var spreadsheet = new Spreadsheet();
        var ex = await Assert.ThrowsAsync<ImportLimitExceededException>(() => spreadsheet.ImportOdsFileAsync(ods));

        Assert.Contains("MaxCells", ex.Message);
    }

    /// <summary>
    /// Repeat counts larger than the sheet limits are clamped to them (1,048,576 rows, 16,384 columns), so an absurd
    /// count can't overflow the indexes, and a row landing exactly on the limit is handled.
    /// </summary>
    [Fact]
    public async Task Ods_RepeatCountsAboveSheetLimits_AreClamped()
    {
        var ods = OdsWithRows(
            "<table:table-row><table:table-cell office:value-type=\"string\" table:number-columns-repeated=\"2147483647\"><text:p>x</text:p></table:table-cell></table:table-row>" +
            "<table:table-row table:number-rows-repeated=\"2147483647\"><table:table-cell/></table:table-row>" +
            "<table:table-row><table:table-cell office:value-type=\"string\"><text:p>past the end</text:p></table:table-cell></table:table-row>");

        using var spreadsheet = new Spreadsheet();
        var sheet = (await spreadsheet.ImportOdsFileAsync(ods)).Sheets[0];

        Assert.Equal(16_384, sheet.Cells.Count(c => c.Row == 1));
        Assert.Equal(16_384, sheet.ColumnCount);
        Assert.Equal("past the end", sheet.Cells.Single(c => c.Row > 1).Value);
    }

    /// <summary>
    /// A package whose entries declare more uncompressed bytes than allowed is rejected before any XML is parsed.
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task DecompressionBomb_StoppedByMaxUncompressedBytes(bool ods)
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        for (int r = 1; r <= 100; r++)
            sheet.AddCell(r, 1, new string('A', 30_000)); // 3 MB of text, within the per-cell limit
        var bytes = ods ? await spreadsheet.GenerateOdsFileAsync() : await spreadsheet.GenerateXlsxFileAsync();
        Assert.True(bytes.Length < 100_000, "expected the repeated text to compress well");

        var options = new ImportOptions { MaxUncompressedBytes = 1_000_000 };
        using var importer = new Spreadsheet();
        var ex = await Assert.ThrowsAsync<ImportLimitExceededException>(() =>
            ods ? importer.ImportOdsFileAsync(bytes, options) : importer.ImportXlsxFileAsync(bytes, options));

        Assert.Contains("MaxUncompressedBytes", ex.Message);

        // The same file imports with the default limit
        using var relaxed = new Spreadsheet();
        var workbook = ods ? await relaxed.ImportOdsFileAsync(bytes) : await relaxed.ImportXlsxFileAsync(bytes);
        Assert.Equal(100, workbook.Sheets[0].Cells.Count);
        Assert.Equal(30_000, workbook.Sheets[0].Cells[0].Value.Length);
    }

    /// <summary>
    /// MaxCells applies to XLSX and delimited files as well.
    /// </summary>
    [Theory]
    [InlineData("xlsx")]
    [InlineData("csv")]
    public async Task MaxCells_AppliesToOtherFormats(string format)
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        for (int r = 1; r <= 10; r++)
            for (int c = 1; c <= 10; c++)
                sheet.AddCell(r, c, "v");
        var bytes = format == "xlsx" ? await spreadsheet.GenerateXlsxFileAsync() : await spreadsheet.GenerateCsvFileAsync();

        var options = new ImportOptions { MaxCells = 50 };
        using var importer = new Spreadsheet();
        await Assert.ThrowsAsync<ImportLimitExceededException>(() =>
            format == "xlsx" ? importer.ImportXlsxFileAsync(bytes, options) : importer.ImportCsvFileAsync(bytes, options));
    }

    /// <summary>
    /// A DTD (the vehicle for entity-expansion and external-entity attacks) is rejected; neither ODF nor OOXML allows one.
    /// </summary>
    [Fact]
    public async Task Ods_Dtd_IsRejected()
    {
        var ods = OdsWithContent(
            "<?xml version=\"1.0\"?><!DOCTYPE x [<!ENTITY a \"aaaaaaaaaa\"><!ENTITY b \"&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;\">]>" +
            $"<office:document-content {Namespaces}><office:body><office:spreadsheet><table:table table:name=\"T\">" +
            "<table:table-row><table:table-cell office:value-type=\"string\"><text:p>&b;</text:p></table:table-cell></table:table-row>" +
            "</table:table></office:spreadsheet></office:body></office:document-content>");

        using var spreadsheet = new Spreadsheet();
        await Assert.ThrowsAsync<System.Xml.XmlException>(() => spreadsheet.ImportOdsFileAsync(ods));
    }

    /// <summary>
    /// A DTD in an XLSX part is rejected the same way.
    /// </summary>
    [Fact]
    public async Task Xlsx_Dtd_IsRejected()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "x");
        var original = await spreadsheet.GenerateXlsxFileAsync();

        using var input = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read);
        using var ms = new MemoryStream();
        using (var output = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            foreach (var entry in input.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                var text = reader.ReadToEnd();
                if (entry.FullName == "xl/worksheets/sheet1.xml")
                    text = text.Replace("?>", "?><!DOCTYPE x [<!ENTITY a \"aaaa\">]>");
                using var stream = output.CreateEntry(entry.FullName).Open();
                var bytes = Encoding.UTF8.GetBytes(text);
                stream.Write(bytes, 0, bytes.Length);
            }
        }

        using var importer = new Spreadsheet();
        await Assert.ThrowsAsync<System.Xml.XmlException>(() => importer.ImportXlsxFileAsync(ms.ToArray()));
    }

    /// <summary>
    /// Row groups and text spans nested tens of thousands deep import without overflowing the stack.
    /// </summary>
    [Fact]
    public async Task Ods_DeeplyNestedMarkup_DoesNotOverflowStack()
    {
        const int depth = 50_000;
        var rows = string.Concat(Enumerable.Repeat("<table:table-row-group>", depth)) +
            "<table:table-row><table:table-cell office:value-type=\"string\"><text:p>" +
            string.Concat(Enumerable.Repeat("<text:span>", depth)) + "deep" + string.Concat(Enumerable.Repeat("</text:span>", depth)) +
            "</text:p></table:table-cell></table:table-row>" +
            string.Concat(Enumerable.Repeat("</table:table-row-group>", depth));

        using var spreadsheet = new Spreadsheet();
        var workbook = await spreadsheet.ImportOdsFileAsync(OdsWithRows(rows));

        Assert.Equal("deep", workbook.Sheets[0].Cells.Single().Value);
    }
}
