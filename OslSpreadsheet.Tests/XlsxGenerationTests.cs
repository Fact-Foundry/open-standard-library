using System.IO.Compression;
using OoxSpreadsheet;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for details of the generated XLSX XML that affect how Excel opens the file.
/// </summary>
public class XlsxGenerationTests
{
    private static string ReadEntry(byte[] package, string path)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry(path)!.Open());
        return reader.ReadToEnd();
    }

    /// <summary>
    /// With several frozen sheets, only the first is marked as the selected tab; otherwise Excel opens them grouped.
    /// </summary>
    [Fact]
    public async Task MultipleFrozenSheets_OnlyFirstIsSelected()
    {
        using var spreadsheet = new Spreadsheet();
        foreach (var name in new[] { "A", "B", "C" })
        {
            var sheet = spreadsheet.Workbook.AddSheet(name);
            sheet.AddCell(1, 1, "Header");
            sheet.FreezeRows = 1;
        }

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.Contains("tabSelected=\"1\"", ReadEntry(bytes, "xl/worksheets/sheet1.xml"));
        Assert.DoesNotContain("tabSelected", ReadEntry(bytes, "xl/worksheets/sheet2.xml"));
        Assert.DoesNotContain("tabSelected", ReadEntry(bytes, "xl/worksheets/sheet3.xml"));
        Assert.Contains("state=\"frozen\"", ReadEntry(bytes, "xl/worksheets/sheet3.xml"));
    }

    /// <summary>
    /// Text cells declare xml:space="preserve" so leading and trailing whitespace is not trimmed by Excel.
    /// </summary>
    [Fact]
    public async Task TextCells_PreserveWhitespace()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "  padded  ");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.Contains("<t xml:space=\"preserve\">  padded  </t>", ReadEntry(bytes, "xl/worksheets/sheet1.xml"));
    }
}
