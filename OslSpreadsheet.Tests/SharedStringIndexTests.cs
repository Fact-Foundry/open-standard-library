using System.IO.Compression;
using System.Text;
using OoxSpreadsheet;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// A shared-string index outside the table yields empty text, whichever direction it's out of range.
/// </summary>
public class SharedStringIndexTests
{
    private static byte[] XlsxWithSharedStrings(params string[] cellIndexes)
    {
        const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string rels = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var cells = new StringBuilder();
        for (int i = 0; i < cellIndexes.Length; i++)
            cells.Append($"<c r=\"{(char)('A' + i)}1\" t=\"s\"><v>{cellIndexes[i]}</v></c>");

        var parts = new (string Path, string Content)[]
        {
            ("[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
                "</Types>"),
            ("_rels/.rels",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{rels}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{relType}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"),
            ("xl/workbook.xml",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"{main}\" xmlns:r=\"{relType}\">" +
                "<sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"),
            ("xl/_rels/workbook.xml.rels",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{rels}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{relType}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                $"<Relationship Id=\"rId2\" Type=\"{relType}/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>"),
            ("xl/sharedStrings.xml",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"{main}\" count=\"2\" uniqueCount=\"2\"><si><t>zero</t></si><si><t>one</t></si></sst>"),
            ("xl/worksheets/sheet1.xml",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"{main}\"><sheetData><row r=\"1\">{cells}</row></sheetData></worksheet>")
        };

        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            foreach (var (path, content) in parts)
            {
                using var stream = archive.CreateEntry(path).Open();
                var bytes = Encoding.UTF8.GetBytes(content);
                stream.Write(bytes, 0, bytes.Length);
            }
        return ms.ToArray();
    }

    [Fact]
    public async Task OutOfRangeIndexes_ImportAsEmptyText()
    {
        var bytes = XlsxWithSharedStrings("0", "1", "2", "-1", "-2147483648", "abc");

        using var spreadsheet = new Spreadsheet();
        var sheet = (await spreadsheet.ImportXlsxFileAsync(bytes)).Sheets[0];

        Assert.Equal("zero", sheet.GetCell(1, 1)?.Value);
        Assert.Equal("one", sheet.GetCell(1, 2)?.Value);
        Assert.Equal("", sheet.GetCell(1, 3)?.Value);   // past the end
        Assert.Equal("", sheet.GetCell(1, 4)?.Value);   // negative
        Assert.Equal("", sheet.GetCell(1, 5)?.Value);   // int.MinValue
        Assert.Equal("zero", sheet.GetCell(1, 6)?.Value); // unparseable falls back to index 0, as before
        Assert.Equal(6, sheet.Cells.Count);
    }
}
