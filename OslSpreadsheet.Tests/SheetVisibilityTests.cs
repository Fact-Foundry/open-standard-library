using System.IO.Compression;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for <see cref="oSpreadsheet.Visibility"/>: reading, writing, and validating hidden sheets in XLSX and ODS.
/// </summary>
public class SheetVisibilityTests
{
    private static string ReadEntry(byte[] package, string path)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var entry = archive.GetEntry(path);
        if (entry == null) return "";
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static Spreadsheet Build(params SheetVisibility[] visibilities)
    {
        var spreadsheet = new Spreadsheet();
        for (int i = 0; i < visibilities.Length; i++)
        {
            var sheet = spreadsheet.Workbook.AddSheet($"S{i + 1}");
            sheet.Visibility = visibilities[i];
            sheet.AddCell(1, 1, $"value {i + 1}");
        }
        return spreadsheet;
    }

    private static byte[] Zip(params (string Path, string Content, CompressionLevel Level)[] entries)
    {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            foreach (var (path, content, level) in entries)
            {
                using var stream = archive.CreateEntry(path, level).Open();
                var bytes = System.Text.Encoding.UTF8.GetBytes(content);
                stream.Write(bytes, 0, bytes.Length);
            }
        }
        return ms.ToArray();
    }

    /// <summary>
    /// A hand-written XLSX shaped like Excel's output: hidden sheets marked only by <c>state</c> on the
    /// workbook's sheet list, a bookViews element, and sheets whose rIds don't follow sheet order.
    /// </summary>
    private static byte[] ExternalXlsx()
    {
        const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string rels = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string worksheetType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

        string Sheet(string text) =>
            $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"{main}\">" +
            $"<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>{text}</t></is></c></row></sheetData></worksheet>";

        var level = CompressionLevel.Fastest;
        return Zip(
            ("[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet4.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "</Types>", level),
            ("_rels/.rels",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{rels}\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>", level),
            ("xl/workbook.xml",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"{main}\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"28800\" windowHeight=\"12300\" activeTab=\"1\"/></bookViews><sheets>" +
                "<sheet name=\"Lookup\" sheetId=\"4\" state=\"hidden\" r:id=\"rId7\"/>" +
                "<sheet name=\"Report\" sheetId=\"1\" r:id=\"rId5\"/>" +
                "<sheet name=\"Secret\" sheetId=\"2\" state=\"veryHidden\" r:id=\"rId6\"/>" +
                "<sheet name=\"Notes\" sheetId=\"3\" state=\"visible\" r:id=\"rId8\"/>" +
                "</sheets></workbook>", level),
            ("xl/_rels/workbook.xml.rels",
                $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{rels}\">" +
                $"<Relationship Id=\"rId5\" Type=\"{worksheetType}\" Target=\"worksheets/sheet1.xml\"/>" +
                $"<Relationship Id=\"rId6\" Type=\"{worksheetType}\" Target=\"worksheets/sheet2.xml\"/>" +
                $"<Relationship Id=\"rId7\" Type=\"{worksheetType}\" Target=\"/xl/worksheets/sheet3.xml\"/>" +
                $"<Relationship Id=\"rId8\" Type=\"{worksheetType}\" Target=\"worksheets/sheet4.xml\"/>" +
                "</Relationships>", level),
            ("xl/worksheets/sheet1.xml", Sheet("report"), level),
            ("xl/worksheets/sheet2.xml", Sheet("secret"), level),
            ("xl/worksheets/sheet3.xml", Sheet("lookup"), level),
            ("xl/worksheets/sheet4.xml", Sheet("notes"), level));
    }

    /// <summary>
    /// A hand-written ODS shaped like LibreOffice's output: table styles named ta1/ta2 shared by several tables,
    /// with <c>table:display</c> on their table-properties, and a table with no style at all.
    /// </summary>
    private static byte[] ExternalOds()
    {
        var content =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
            "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" " +
            "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" " +
            "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\">" +
            "<office:automatic-styles>" +
            "<style:style style:name=\"co1\" style:family=\"table-column\"><style:table-column-properties style:column-width=\"2.258cm\"/></style:style>" +
            "<style:style style:name=\"ta1\" style:family=\"table\" style:master-page-name=\"Default\"><style:table-properties table:display=\"true\" style:writing-mode=\"lr-tb\"/></style:style>" +
            "<style:style style:name=\"ta2\" style:family=\"table\" style:master-page-name=\"Default\"><style:table-properties table:display=\"false\" style:writing-mode=\"lr-tb\"/></style:style>" +
            "</office:automatic-styles>" +
            "<office:body><office:spreadsheet>" +
            Table("Lookup", "ta2", "lookup") +
            Table("Report", "ta1", "report") +
            Table("Archive", "ta2", "archive") +
            Table("Unstyled", null, "unstyled") +
            "</office:spreadsheet></office:body></office:document-content>";

        static string Table(string name, string? style, string text) =>
            $"<table:table table:name=\"{name}\"{(style == null ? "" : $" table:style-name=\"{style}\"")}>" +
            "<table:table-column table:style-name=\"co1\"/>" +
            $"<table:table-row><table:table-cell office:value-type=\"string\"><text:p>{text}</text:p></table:table-cell></table:table-row>" +
            "</table:table>";

        return Zip(
            ("mimetype", "application/vnd.oasis.opendocument.spreadsheet", CompressionLevel.NoCompression),
            ("META-INF/manifest.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\">" +
                "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>" +
                "<manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>" +
                "</manifest:manifest>", CompressionLevel.Fastest),
            ("content.xml", content, CompressionLevel.Fastest));
    }

    [Fact]
    public void Visibility_DefaultsToVisible()
    {
        Assert.Equal(SheetVisibility.Visible, new oWorkbook().AddSheet().Visibility);
    }

    // ----- XLSX -----

    [Fact]
    public async Task Xlsx_WritesState()
    {
        using var spreadsheet = Build(SheetVisibility.Visible, SheetVisibility.Hidden, SheetVisibility.VeryHidden);

        var workbookXml = ReadEntry(await spreadsheet.GenerateXlsxFileAsync(), "xl/workbook.xml");

        Assert.Contains("<sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/>", workbookXml);
        Assert.Contains("<sheet name=\"S2\" sheetId=\"2\" state=\"hidden\" r:id=\"rId2\"/>", workbookXml);
        Assert.Contains("<sheet name=\"S3\" sheetId=\"3\" state=\"veryHidden\" r:id=\"rId3\"/>", workbookXml);
        Assert.DoesNotContain("bookViews", workbookXml);
    }

    [Fact]
    public async Task Xlsx_RoundTripsVisibility()
    {
        using var source = Build(SheetVisibility.Visible, SheetVisibility.Hidden, SheetVisibility.VeryHidden);
        var bytes = await source.GenerateXlsxFileAsync();

        using var spreadsheet = new Spreadsheet();
        var workbook = await spreadsheet.ImportXlsxFileAsync(bytes);

        Assert.Equal(
            new[] { SheetVisibility.Visible, SheetVisibility.Hidden, SheetVisibility.VeryHidden },
            workbook.Sheets.Select(s => s.Visibility));
    }

    [Fact]
    public async Task Xlsx_ImportsStateFromExternalFile()
    {
        var bytes = ExternalXlsx();
        Assert.True(SpreadsheetValidator.ValidateXlsx(bytes).IsValid);

        using var spreadsheet = new Spreadsheet();
        var workbook = await spreadsheet.ImportXlsxFileAsync(bytes);

        Assert.Equal(new[] { "Lookup", "Report", "Secret", "Notes" }, workbook.Sheets.Select(s => s.SheetName));
        Assert.Equal(
            new[] { SheetVisibility.Hidden, SheetVisibility.Visible, SheetVisibility.VeryHidden, SheetVisibility.Visible },
            workbook.Sheets.Select(s => s.Visibility));
        Assert.Equal(new[] { "lookup", "report", "secret", "notes" }, workbook.Sheets.Select(s => s.GetCell(1, 1)?.Value));
    }

    [Fact]
    public async Task Xlsx_HiddenFirstSheet_OpensOnFirstVisible()
    {
        using var spreadsheet = Build(SheetVisibility.Hidden, SheetVisibility.VeryHidden, SheetVisibility.Visible);
        spreadsheet.Workbook.Sheets[0].FreezeRows = 1;
        spreadsheet.Workbook.Sheets[2].FreezeRows = 1;

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.Contains("<bookViews><workbookView activeTab=\"2\"/></bookViews><sheets>", ReadEntry(bytes, "xl/workbook.xml"));
        Assert.DoesNotContain("tabSelected", ReadEntry(bytes, "xl/worksheets/sheet1.xml"));
        Assert.Contains("tabSelected=\"1\"", ReadEntry(bytes, "xl/worksheets/sheet3.xml"));
    }

    [Fact]
    public async Task Xlsx_GeneratedFileWithHiddenSheetsIsValid()
    {
        using var spreadsheet = Build(SheetVisibility.Hidden, SheetVisibility.Visible, SheetVisibility.VeryHidden);

        var result = SpreadsheetValidator.ValidateXlsx(await spreadsheet.GenerateXlsxFileAsync());

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(i => i.Code)));
        Assert.DoesNotContain(result.Issues, i => i.Code == "XLSX_ACTIVE_TAB_INVALID");
    }

    [Fact]
    public async Task Xlsx_AllHidden_GenerateThrows()
    {
        using var spreadsheet = Build(SheetVisibility.Hidden, SheetVisibility.VeryHidden);

        var ex = await Assert.ThrowsAsync<InvalidWorkbookException>(() => spreadsheet.GenerateXlsxFileAsync());

        Assert.Contains(ex.Result.Issues, i => i.Code == "WB_NO_VISIBLE_SHEET" && i.Severity == ValidationSeverity.Error);
    }

    // ----- ODS -----

    [Fact]
    public async Task Ods_WritesDisplayOnTableStyle()
    {
        using var spreadsheet = Build(SheetVisibility.Visible, SheetVisibility.Hidden, SheetVisibility.VeryHidden);

        var content = ReadEntry(await spreadsheet.GenerateOdsFileAsync(), "content.xml");

        Assert.Matches("style:name=\"ta1\"[^>]*>\\s*<style:table-properties table:display=\"true\"", content);
        Assert.Matches("style:name=\"ta2\"[^>]*>\\s*<style:table-properties table:display=\"false\"", content);
        Assert.Matches("style:name=\"ta3\"[^>]*>\\s*<style:table-properties table:display=\"false\"", content);
    }

    [Fact]
    public async Task Ods_RoundTripsVisibility_VeryHiddenBecomesHidden()
    {
        using var source = Build(SheetVisibility.Visible, SheetVisibility.Hidden, SheetVisibility.VeryHidden);
        var bytes = await source.GenerateOdsFileAsync();

        using var spreadsheet = new Spreadsheet();
        var workbook = await spreadsheet.ImportOdsFileAsync(bytes);

        Assert.Equal(
            new[] { SheetVisibility.Visible, SheetVisibility.Hidden, SheetVisibility.Hidden },
            workbook.Sheets.Select(s => s.Visibility));
    }

    [Fact]
    public async Task Ods_ImportsDisplayFromExternalFile()
    {
        var bytes = ExternalOds();
        var validation = SpreadsheetValidator.ValidateOds(bytes);
        Assert.True(validation.IsValid, string.Join("; ", validation.Issues.Select(i => i.Code)));
        Assert.DoesNotContain(validation.Issues, i => i.Code == "ODS_NO_VISIBLE_SHEET");

        using var spreadsheet = new Spreadsheet();
        var workbook = await spreadsheet.ImportOdsFileAsync(bytes);

        Assert.Equal(new[] { "Lookup", "Report", "Archive", "Unstyled" }, workbook.Sheets.Select(s => s.SheetName));
        Assert.Equal(
            new[] { SheetVisibility.Hidden, SheetVisibility.Visible, SheetVisibility.Hidden, SheetVisibility.Visible },
            workbook.Sheets.Select(s => s.Visibility));
        Assert.Equal(new[] { "lookup", "report", "archive", "unstyled" }, workbook.Sheets.Select(s => s.GetCell(1, 1)?.Value));
    }

    [Fact]
    public async Task Ods_HiddenFirstSheet_OpensOnFirstVisible()
    {
        using var spreadsheet = Build(SheetVisibility.Hidden, SheetVisibility.Visible);

        var settings = ReadEntry(await spreadsheet.GenerateOdsFileAsync(), "settings.xml");

        Assert.Contains("<config:config-item config:name=\"ActiveTable\" config:type=\"string\">S2</config:config-item>", settings);
    }

    [Fact]
    public async Task Ods_VisibleFirstSheet_WritesNoSettings()
    {
        using var spreadsheet = Build(SheetVisibility.Visible, SheetVisibility.Hidden);

        Assert.Equal("", ReadEntry(await spreadsheet.GenerateOdsFileAsync(), "settings.xml"));
    }

    [Fact]
    public async Task Ods_GeneratedFileWithHiddenSheetsIsValid()
    {
        using var spreadsheet = Build(SheetVisibility.Hidden, SheetVisibility.Visible);

        var result = SpreadsheetValidator.ValidateOds(await spreadsheet.GenerateOdsFileAsync());

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(i => i.Code)));
        Assert.DoesNotContain(result.Issues, i => i.Code == "ODS_NO_VISIBLE_SHEET");
    }

    [Fact]
    public async Task Ods_AllHidden_GeneratesWithWarnings()
    {
        using var spreadsheet = Build(SheetVisibility.Hidden, SheetVisibility.Hidden);

        Assert.Contains(spreadsheet.Workbook.Validate(ValidationFileFormat.Ods).Issues,
            i => i.Code == "WB_NO_VISIBLE_SHEET" && i.Severity == ValidationSeverity.Warning);

        var result = SpreadsheetValidator.ValidateOds(await spreadsheet.GenerateOdsFileAsync());

        Assert.True(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == "ODS_NO_VISIBLE_SHEET" && i.Severity == ValidationSeverity.Warning);
    }

    [Fact]
    public void Ods_VeryHidden_IsWarning()
    {
        using var spreadsheet = Build(SheetVisibility.Visible, SheetVisibility.VeryHidden);

        var issues = spreadsheet.Workbook.Validate(ValidationFileFormat.Ods).Issues;

        Assert.Contains(issues, i => i.Code == "WB_SHEET_VERY_HIDDEN_UNSUPPORTED" && i.Severity == ValidationSeverity.Warning && i.Sheet == "S2");
        Assert.DoesNotContain(spreadsheet.Workbook.Validate(ValidationFileFormat.Xlsx).Issues, i => i.Code == "WB_SHEET_VERY_HIDDEN_UNSUPPORTED");
    }

    // ----- Workbook validation and delimited -----

    [Fact]
    public void UndefinedVisibility_IsError()
    {
        using var spreadsheet = Build(SheetVisibility.Visible, (SheetVisibility)7);

        Assert.Contains(spreadsheet.Workbook.Validate(ValidationFileFormat.Xlsx).Issues,
            i => i.Code == "WB_SHEET_VISIBILITY_INVALID" && i.Severity == ValidationSeverity.Error);
    }

    [Fact]
    public async Task Delimited_IgnoresVisibility()
    {
        using var spreadsheet = Build(SheetVisibility.Hidden);

        Assert.DoesNotContain(spreadsheet.Workbook.Validate(ValidationFileFormat.Delimited).Issues, i => i.Code.Contains("VISIBLE"));
        Assert.Equal("\"value 1\"", System.Text.Encoding.UTF8.GetString(await spreadsheet.GenerateCsvFileAsync()).TrimEnd());
    }
}
