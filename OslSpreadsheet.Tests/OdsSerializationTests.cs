using System.IO.Compression;
using System.Text;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// The ODS parts are written exactly as serialized: no text substitution that could reach into cell data,
/// and no swallowed errors that could leave an empty part in the package.
/// </summary>
public class OdsSerializationTests
{
    private static string ReadEntry(byte[] package, string path)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry(path)!.Open());
        return reader.ReadToEnd();
    }

    /// <summary>
    /// Text that happens to match what the old post-processing replaced must survive a round trip unchanged.
    /// </summary>
    [Theory]
    [InlineData("encoding is utf-8 here")]
    [InlineData("utf-8")]
    [InlineData("<img width=\"5\" />")]
    [InlineData("a\" />b")]
    [InlineData("UTF-8 and utf-8 and \" /> all at once")]
    public async Task CellText_RoundTripsUnchanged(string text)
    {
        using var source = new Spreadsheet();
        source.Workbook.AddSheet("Data").AddCell(1, 1, text);
        var bytes = await source.GenerateOdsFileAsync();

        Assert.True(SpreadsheetValidator.ValidateOds(bytes).IsValid);

        using var spreadsheet = new Spreadsheet();
        var imported = await spreadsheet.ImportOdsFileAsync(bytes);

        Assert.Equal(text, imported.Sheets[0].GetCell(1, 1)?.Value);
    }

    [Fact]
    public async Task SheetNameAndMetadata_RoundTripUnchanged()
    {
        using var source = new Spreadsheet();
        source.Workbook.Creator = "utf-8 \" /> tester";
        source.Workbook.AddSheet("Sheet utf-8").AddCell(1, 1, "x");
        var bytes = await source.GenerateOdsFileAsync();

        Assert.Contains("Sheet utf-8", ReadEntry(bytes, "content.xml"));
        Assert.Contains("utf-8 \" /&gt; tester", ReadEntry(bytes, "meta.xml")); // '>' is escaped in element text, '"' is not

        using var spreadsheet = new Spreadsheet();
        Assert.Equal("Sheet utf-8", (await spreadsheet.ImportOdsFileAsync(bytes)).Sheets[0].SheetName);
    }

    [Fact]
    public async Task EveryPart_IsNonEmptyWellFormedXml()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "x");
        sheet.FreezeRows = 1;
        var bytes = await spreadsheet.GenerateOdsFileAsync();

        foreach (var part in new[] { "META-INF/manifest.xml", "content.xml", "meta.xml", "styles.xml", "settings.xml" })
        {
            var xml = ReadEntry(bytes, part);
            Assert.False(string.IsNullOrWhiteSpace(xml), $"{part} is empty");
            Assert.StartsWith("<?xml version=\"1.0\" encoding=\"utf-8\"?>", xml, StringComparison.OrdinalIgnoreCase);
            System.Xml.Linq.XDocument.Parse(xml); // throws if not well-formed
        }
    }

    [Fact]
    public async Task GeneratedOds_OpensAsUtf8()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "héllo €");
        var bytes = await spreadsheet.GenerateOdsFileAsync();

        var content = ReadEntry(bytes, "content.xml");
        Assert.Contains("héllo €", content);
        Assert.Equal(content, Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(content)));
        Assert.True(SpreadsheetValidator.ValidateOds(bytes).IsValid);
    }
}
