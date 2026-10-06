using System.IO.Compression;
using System.Text;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for <see cref="SpreadsheetValidator"/>: files produced by the library must validate cleanly,
/// and common malformations must be reported with the expected code, severity, and location.
/// </summary>
public class ValidationTests
{
    // --- Helpers ---

    /// <summary>
    /// Builds a workbook exercising styles, dates, booleans, freeze panes, auto-filters, and multiple sheets.
    /// </summary>
    private static Spreadsheet BuildFullFeaturedSpreadsheet()
    {
        var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Score");
        sheet.AddCell(1, 3, "Joined");
        sheet.AddCell(1, 4, "Active");
        sheet.AddCell(2, 1, "Alice & \"Al\" <A>");
        sheet.AddCell(2, 2, "95.5", CellValueType.Float);
        sheet.AddCell(2, 3, "2024-01-31", CellValueType.DateTime);
        sheet.AddCell(2, 4, "true", CellValueType.Boolean);
        sheet.AddCell(3, 1, "Bob");
        sheet.AddCell(3, 2, "82", CellValueType.Float);
        sheet.AddCell(3, 3, "2024-02-01T13:45:00", CellValueType.DateTime);
        sheet.AddCell(3, 4, "false", CellValueType.Boolean);
        sheet.Cells[0].Style = new CellStyle { Bold = true, BackgroundColor = "#FFFF00" };
        sheet.FreezeRows = 1;
        sheet.SetAutoFilter();

        var other = spreadsheet.Workbook.AddSheet("Other");
        other.AddCell(1, 1, "x");
        return spreadsheet;
    }

    /// <summary>
    /// Builds a ZIP archive from (path, content) pairs, in order, with every entry deflated.
    /// </summary>
    private static byte[] Zip(params (string Path, string Content)[] entries)
    {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            foreach (var (path, content) in entries)
            {
                using var stream = archive.CreateEntry(path).Open();
                var bytes = Encoding.UTF8.GetBytes(content);
                stream.Write(bytes, 0, bytes.Length);
            }
        }
        return ms.ToArray();
    }

    private const string ContentTypes =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
        "</Types>";

    private const string RootRels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
        "</Relationships>";

    private const string WorkbookXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>" +
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";

    private const string WorkbookRels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
        "</Relationships>";

    private const string ValidSheetData =
        "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\"><v>42</v></c></row></sheetData>";

    /// <summary>
    /// Builds a hand-written minimal XLSX. Any part can be replaced to introduce a specific defect.
    /// </summary>
    private static byte[] MinimalXlsx(
        string? sheetBody = null,
        string contentTypes = ContentTypes,
        string workbook = WorkbookXml,
        string workbookRels = WorkbookRels,
        params (string Path, string Content)[] extraParts)
    {
        var sheet = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    (sheetBody ?? ValidSheetData) + "</worksheet>";

        var parts = new List<(string, string)>
        {
            ("[Content_Types].xml", contentTypes),
            ("_rels/.rels", RootRels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbookRels),
            ("xl/worksheets/sheet1.xml", sheet)
        };
        parts.AddRange(extraParts);
        return Zip(parts.ToArray());
    }

    /// <summary>
    /// Rewrites one entry of an existing ZIP package, preserving entry order and the stored mimetype entry.
    /// </summary>
    private static byte[] ReplaceEntry(byte[] package, string path, Func<string, string> transform)
    {
        using var input = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var ms = new MemoryStream();
        using (var output = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            foreach (var entry in input.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                var content = reader.ReadToEnd();
                if (entry.FullName == path)
                    content = transform(content);

                var level = entry.FullName == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Fastest;
                using var stream = output.CreateEntry(entry.FullName, level).Open();
                var bytes = Encoding.UTF8.GetBytes(content);
                stream.Write(bytes, 0, bytes.Length);
            }
        }
        return ms.ToArray();
    }

    private static ValidationIssue AssertHasIssue(ValidationResult result, string code, ValidationSeverity severity)
    {
        var issue = result.Issues.FirstOrDefault(i => i.Code == code);
        Assert.True(issue != null, $"Expected issue {code}, but got:\n{result}");
        Assert.Equal(severity, issue!.Severity);
        return issue;
    }

    private static void AssertNoIssues(ValidationResult result) =>
        Assert.True(result.Issues.Count == 0, result.ToString());

    // --- Library-generated files ---

    /// <summary>
    /// XLSX files generated by this library must pass validation with no errors or warnings.
    /// </summary>
    [Fact]
    public async Task Xlsx_GeneratedByLibrary_IsValid()
    {
        using var spreadsheet = BuildFullFeaturedSpreadsheet();
        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        var result = SpreadsheetValidator.ValidateXlsx(bytes);

        Assert.True(result.IsValid, result.ToString());
        AssertNoIssues(result);
    }

    /// <summary>
    /// ODS files generated by this library must pass validation with no errors or warnings.
    /// </summary>
    [Fact]
    public async Task Ods_GeneratedByLibrary_IsValid()
    {
        using var spreadsheet = BuildFullFeaturedSpreadsheet();
        var bytes = await spreadsheet.GenerateOdsFileAsync();

        var result = SpreadsheetValidator.ValidateOds(bytes);

        Assert.True(result.IsValid, result.ToString());
        AssertNoIssues(result);
    }

    /// <summary>
    /// CSV, tab, pipe, and ASCII-delimited files generated by this library must pass validation.
    /// </summary>
    [Theory]
    [InlineData(ColumnDelimeter.Comma)]
    [InlineData(ColumnDelimeter.Tab)]
    [InlineData(ColumnDelimeter.Pipe)]
    [InlineData(ColumnDelimeter.ASCII)]
    public async Task Delimited_GeneratedByLibrary_IsValid(ColumnDelimeter delimiter)
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.ColumnDelimeter = delimiter;
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Quote");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, delimiter == ColumnDelimeter.Comma ? "She said \"hi\", then left" : "plain");
        var bytes = await spreadsheet.GenerateCsvFileAsync();

        var result = SpreadsheetValidator.ValidateDelimited(bytes, delimiter);

        Assert.True(result.IsValid, result.ToString());
        AssertNoIssues(result);
    }

    /// <summary>
    /// The hand-written minimal XLSX used by the negative tests must itself be valid.
    /// </summary>
    [Fact]
    public void Xlsx_MinimalHandWritten_IsValid()
    {
        var result = SpreadsheetValidator.ValidateXlsx(MinimalXlsx());
        AssertNoIssues(result);
    }

    // --- Format detection and routing ---

    /// <summary>
    /// DetectFormat identifies each format from content alone.
    /// </summary>
    [Fact]
    public async Task DetectFormat_IdentifiesFormatsFromContent()
    {
        using var spreadsheet = BuildFullFeaturedSpreadsheet();

        Assert.Equal(ValidationFileFormat.Xlsx, SpreadsheetValidator.DetectFormat(await spreadsheet.GenerateXlsxFileAsync()));
        Assert.Equal(ValidationFileFormat.Ods, SpreadsheetValidator.DetectFormat(await spreadsheet.GenerateOdsFileAsync()));
        Assert.Equal(ValidationFileFormat.Delimited, SpreadsheetValidator.DetectFormat("a,b\n1,2\n"u8.ToArray()));
        Assert.Equal(ValidationFileFormat.Unknown, SpreadsheetValidator.DetectFormat("%PDF-1.7"u8.ToArray()));
        Assert.Equal(ValidationFileFormat.Unknown, SpreadsheetValidator.DetectFormat("{\"rows\": []}"u8.ToArray()));
        Assert.Equal(ValidationFileFormat.Unknown, SpreadsheetValidator.DetectFormat([]));
    }

    /// <summary>
    /// A CSV saved with an .xlsx extension is rejected with a message that says what the file actually is.
    /// </summary>
    [Fact]
    public void Validate_CsvTextNamedXlsx_ReportsNotZip()
    {
        var result = SpreadsheetValidator.Validate("Name,Score\nAlice,95\n"u8.ToArray(), "report.xlsx");

        Assert.False(result.IsValid);
        Assert.Equal(ValidationFileFormat.Xlsx, result.Format);
        var issue = AssertHasIssue(result, "XLSX_NOT_ZIP", ValidationSeverity.Error);
        Assert.Contains("plain text", issue.Message);
    }

    /// <summary>
    /// Base64 text of a ZIP (a common LLM output mistake) is identified as such.
    /// </summary>
    [Fact]
    public void Validate_Base64EncodedXlsx_ReportsBase64()
    {
        var base64 = Encoding.ASCII.GetBytes(Convert.ToBase64String(MinimalXlsx()));

        var result = SpreadsheetValidator.Validate(base64, "report.xlsx");

        Assert.Contains("base64", AssertHasIssue(result, "XLSX_NOT_ZIP", ValidationSeverity.Error).Message);
    }

    /// <summary>
    /// An ODS file named .xlsx is identified as ODS.
    /// </summary>
    [Fact]
    public async Task Validate_OdsNamedXlsx_ReportsOds()
    {
        using var spreadsheet = BuildFullFeaturedSpreadsheet();
        var ods = await spreadsheet.GenerateOdsFileAsync();

        var result = SpreadsheetValidator.Validate(ods, "report.xlsx");

        AssertHasIssue(result, "XLSX_IS_ODS", ValidationSeverity.Error);
    }

    /// <summary>
    /// The .tsv extension selects the tab delimiter.
    /// </summary>
    [Fact]
    public void Validate_TsvExtension_UsesTabDelimiter()
    {
        var result = SpreadsheetValidator.Validate("a\tb\n1\t2\n"u8.ToArray(), "data.tsv");

        Assert.Equal(ValidationFileFormat.Delimited, result.Format);
        AssertNoIssues(result);
    }

    /// <summary>
    /// Unrecognized extensions fall back to content detection, and unrecognized content is reported.
    /// </summary>
    [Fact]
    public void Validate_UnknownExtensionAndContent_ReportsUnrecognized()
    {
        var result = SpreadsheetValidator.Validate("<html><body>table</body></html>"u8.ToArray(), "report.dat");

        Assert.Contains("HTML", AssertHasIssue(result, "FORMAT_UNRECOGNIZED", ValidationSeverity.Error).Message);
    }

    // --- XLSX defects ---

    /// <summary>
    /// A missing [Content_Types].xml makes the package invalid.
    /// </summary>
    [Fact]
    public void Xlsx_MissingContentTypes_IsError()
    {
        var bytes = Zip(("_rels/.rels", RootRels), ("xl/workbook.xml", WorkbookXml));
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_CONTENT_TYPES_MISSING", ValidationSeverity.Error);
    }

    /// <summary>
    /// Malformed XML is reported with the part and line.
    /// </summary>
    [Fact]
    public void Xlsx_UnescapedAmpersand_IsMalformedXml()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Fish & Chips</t></is></c></row></sheetData>");

        var issue = AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_XML_MALFORMED", ValidationSeverity.Error);
        Assert.Equal("xl/worksheets/sheet1.xml", issue.Part);
        Assert.Contains("line 1", issue.Location);
    }

    /// <summary>
    /// Worksheet elements out of schema order (cols after sheetData) are an error.
    /// </summary>
    [Fact]
    public void Xlsx_ColsAfterSheetData_IsOrderError()
    {
        var bytes = MinimalXlsx(ValidSheetData + "<cols><col min=\"1\" max=\"1\" width=\"20\" customWidth=\"1\"/></cols>");

        var issue = AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_ELEMENT_ORDER", ValidationSeverity.Error);
        Assert.Contains("<cols> must come before <sheetData>", issue.Message);
    }

    /// <summary>
    /// Text stored in a numeric cell is an error with structured sheet, cell, and row fields.
    /// </summary>
    [Fact]
    public void Xlsx_TextInNumericCell_ReportsStructuredLocation()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"3\"><c r=\"B3\"><v>Alice</v></c></row></sheetData>");

        var issue = AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_CELL_NOT_NUMERIC", ValidationSeverity.Error);
        Assert.Equal("Sheet1", issue.Sheet);
        Assert.Equal("B3", issue.Cell);
        Assert.Equal(3, issue.Row);
        Assert.Equal("xl/worksheets/sheet1.xml", issue.Part);
    }

    /// <summary>
    /// t="inlineStr" with a v element instead of is is an error.
    /// </summary>
    [Fact]
    public void Xlsx_InlineStringWithValueElement_IsError()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><v>Name</v></c></row></sheetData>");
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_INLINE_STRING_MISSING", ValidationSeverity.Error);
    }

    /// <summary>
    /// A shared string reference without a shared strings part is an error.
    /// </summary>
    [Fact]
    public void Xlsx_SharedStringWithoutSharedStringsPart_IsError()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData>");
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_SHARED_STRINGS_MISSING", ValidationSeverity.Error);
    }

    /// <summary>
    /// Rows out of ascending order are an error.
    /// </summary>
    [Fact]
    public void Xlsx_RowsOutOfOrder_IsError()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"2\"><c r=\"A2\"><v>1</v></c></row><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>");
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_ROW_ORDER", ValidationSeverity.Error);
    }

    /// <summary>
    /// A cell whose reference names a different row than its parent row is an error.
    /// </summary>
    [Fact]
    public void Xlsx_CellRowMismatch_IsError()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"1\"><c r=\"A2\"><v>1</v></c></row></sheetData>");
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_CELL_ROW_MISMATCH", ValidationSeverity.Error);
    }

    /// <summary>
    /// A relationship whose target part does not exist is an error.
    /// </summary>
    [Fact]
    public void Xlsx_RelationshipTargetMissing_IsError()
    {
        var rels = WorkbookRels.Replace("worksheets/sheet1.xml", "worksheets/sheet9.xml");
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(MinimalXlsx(workbookRels: rels)), "XLSX_RELATIONSHIP_TARGET_MISSING", ValidationSeverity.Error);
    }

    /// <summary>
    /// Sheet names longer than 31 characters or containing reserved characters are errors.
    /// </summary>
    [Theory]
    [InlineData("Quarterly Revenue Summary For 2024")]
    [InlineData("Q1/Q2")]
    [InlineData("Data[1]")]
    public void Xlsx_InvalidSheetName_IsError(string name)
    {
        var workbook = WorkbookXml.Replace("name=\"Sheet1\"", $"name=\"{name}\"");
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(MinimalXlsx(workbook: workbook)), "XLSX_SHEET_NAME_INVALID", ValidationSeverity.Error);
    }

    /// <summary>
    /// A cell style index without a styles part is an error.
    /// </summary>
    [Fact]
    public void Xlsx_StyleIndexWithoutStyles_IsError()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"1\"><c r=\"A1\" s=\"3\"><v>1</v></c></row></sheetData>");
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_CELL_STYLE_OUT_OF_RANGE", ValidationSeverity.Error);
    }

    /// <summary>
    /// Overlapping merged ranges are an error.
    /// </summary>
    [Fact]
    public void Xlsx_OverlappingMerges_IsError()
    {
        var bytes = MinimalXlsx(ValidSheetData + "<mergeCells count=\"2\"><mergeCell ref=\"A1:B2\"/><mergeCell ref=\"B2:C3\"/></mergeCells>");
        AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_MERGE_OVERLAP", ValidationSeverity.Error);
    }

    /// <summary>
    /// A formula containing #REF! is an error.
    /// </summary>
    [Fact]
    public void Xlsx_FormulaWithBrokenReference_IsError()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"1\"><c r=\"A1\"><f>SUM(#REF!)</f></c></row></sheetData>");
        var issue = AssertHasIssue(SpreadsheetValidator.ValidateXlsx(bytes), "XLSX_FORMULA_BROKEN_REFERENCE", ValidationSeverity.Error);
        Assert.Equal("A1", issue.Cell);
    }

    /// <summary>
    /// A cached #REF! result is an error, but #N/A is only a warning because it can be a legitimate result.
    /// </summary>
    [Fact]
    public void Xlsx_CachedErrorValues_SeverityDependsOnError()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"1\"><c r=\"A1\" t=\"e\"><f>VLOOKUP(1,B:B,1,0)</f><v>#N/A</v></c>" +
                                "<c r=\"B1\" t=\"e\"><f>C1</f><v>#REF!</v></c></row></sheetData>");

        var result = SpreadsheetValidator.ValidateXlsx(bytes);

        Assert.Contains(result.Issues, i => i.Code == "XLSX_CELL_FORMULA_ERROR" && i.Cell == "A1" && i.Severity == ValidationSeverity.Warning);
        Assert.Contains(result.Issues, i => i.Code == "XLSX_CELL_FORMULA_ERROR" && i.Cell == "B1" && i.Severity == ValidationSeverity.Error);
    }

    /// <summary>
    /// An empty worksheet is a warning and does not make the file invalid.
    /// </summary>
    [Fact]
    public void Xlsx_EmptySheet_IsWarningOnly()
    {
        var result = SpreadsheetValidator.ValidateXlsx(MinimalXlsx("<sheetData/>"));

        Assert.True(result.IsValid, result.ToString());
        Assert.Equal("Sheet1", AssertHasIssue(result, "XLSX_SHEET_EMPTY", ValidationSeverity.Warning).Sheet);
    }

    /// <summary>
    /// Validation stops at MaxIssues and marks the result as truncated.
    /// </summary>
    [Fact]
    public void Xlsx_ManyIssues_TruncatesAtMaxIssues()
    {
        var rows = string.Concat(Enumerable.Range(1, 50).Select(r => $"<row r=\"{r}\"><c r=\"A{r}\"><v>text</v></c></row>"));
        var result = SpreadsheetValidator.ValidateXlsx(MinimalXlsx($"<sheetData>{rows}</sheetData>"), new ValidationOptions { MaxIssues = 5 });

        Assert.Equal(5, result.Issues.Count);
        Assert.True(result.Truncated);
        Assert.Contains("stopped after the maximum number of issues", result.ToString());
    }

    // --- ODS defects ---

    /// <summary>
    /// A float cell without office:value is an error located by sheet and cell.
    /// </summary>
    [Fact]
    public async Task Ods_FloatCellWithoutValue_IsError()
    {
        using var spreadsheet = BuildFullFeaturedSpreadsheet();
        var bytes = ReplaceEntry(await spreadsheet.GenerateOdsFileAsync(), "content.xml",
            xml => xml.Replace("office:value=\"95.5\"", ""));

        var issue = AssertHasIssue(SpreadsheetValidator.ValidateOds(bytes), "ODS_VALUE_MISSING", ValidationSeverity.Error);
        Assert.Equal("Data", issue.Sheet);
        Assert.Equal("B2", issue.Cell);
        Assert.Equal(2, issue.Row);
    }

    /// <summary>
    /// A file in the archive that is not listed in the manifest is an error (LibreOffice treats it as corruption).
    /// </summary>
    [Fact]
    public async Task Ods_FileNotInManifest_IsError()
    {
        using var spreadsheet = BuildFullFeaturedSpreadsheet();
        var bytes = ReplaceEntry(await spreadsheet.GenerateOdsFileAsync(), "META-INF/manifest.xml",
            xml => System.Text.RegularExpressions.Regex.Replace(xml, "<manifest:file-entry manifest:full-path=\"styles.xml\"[^>]*/>", ""));

        AssertHasIssue(SpreadsheetValidator.ValidateOds(bytes), "ODS_MANIFEST_FILE_UNLISTED", ValidationSeverity.Error);
    }

    /// <summary>
    /// An XLSX file validated as ODS is identified as XLSX.
    /// </summary>
    [Fact]
    public async Task Ods_GivenXlsx_ReportsXlsx()
    {
        using var spreadsheet = BuildFullFeaturedSpreadsheet();
        AssertHasIssue(SpreadsheetValidator.ValidateOds(await spreadsheet.GenerateXlsxFileAsync()), "ODS_IS_XLSX", ValidationSeverity.Error);
    }

    // --- Delimited defects ---

    /// <summary>
    /// RFC 4180 features (quoted delimiters, doubled quotes, line breaks in quoted values, CRLF) are valid.
    /// </summary>
    [Fact]
    public void Csv_Rfc4180Features_AreValid()
    {
        var csv = "Name,Quote,Notes\r\n\"Smith, John\",\"He said \"\"hi\"\"\",\"line one\r\nline two\"\r\nJane,plain,\r\n";
        AssertNoIssues(SpreadsheetValidator.ValidateDelimited(Encoding.UTF8.GetBytes(csv)));
    }

    /// <summary>
    /// An unclosed quote is an error that points at the line where the quote started.
    /// </summary>
    [Fact]
    public void Csv_UnterminatedQuote_IsError()
    {
        var issue = AssertHasIssue(SpreadsheetValidator.ValidateDelimited("a,b\n1,\"oops\n2,3\n"u8.ToArray()),
            "CSV_UNTERMINATED_QUOTE", ValidationSeverity.Error);
        Assert.Contains("line 2", issue.Location);
    }

    /// <summary>
    /// Undoubled quotes inside a quoted value are an error.
    /// </summary>
    [Fact]
    public void Csv_UndoubledInnerQuotes_IsError()
    {
        var result = SpreadsheetValidator.ValidateDelimited("a,b\n1,\"He said \"hi\" today\"\n"u8.ToArray());
        AssertHasIssue(result, "CSV_TEXT_AFTER_CLOSING_QUOTE", ValidationSeverity.Error);
    }

    /// <summary>
    /// A row with a different field count is an error with the row number; it is a warning when consistency isn't required.
    /// </summary>
    [Fact]
    public void Csv_FieldCountMismatch_IsErrorUnlessRelaxed()
    {
        var csv = "Name,City\nAlice,Paris\nBob,Portland, OR\n"u8.ToArray();

        var strict = AssertHasIssue(SpreadsheetValidator.ValidateDelimited(csv), "CSV_FIELD_COUNT_MISMATCH", ValidationSeverity.Error);
        Assert.Equal(3, strict.Row);

        var relaxed = SpreadsheetValidator.ValidateDelimited(csv, new ValidationOptions { RequireConsistentColumnCount = false });
        Assert.True(relaxed.IsValid);
        AssertHasIssue(relaxed, "CSV_FIELD_COUNT_MISMATCH", ValidationSeverity.Warning);
    }

    /// <summary>
    /// A quote inside an unquoted value is a warning.
    /// </summary>
    [Fact]
    public void Csv_BareQuote_IsWarning()
    {
        var result = SpreadsheetValidator.ValidateDelimited("Item,Size\nTV,55\" screen\n"u8.ToArray());

        Assert.True(result.IsValid, result.ToString());
        AssertHasIssue(result, "CSV_BARE_QUOTE", ValidationSeverity.Warning);
    }

    /// <summary>
    /// A tab-delimited file validated as CSV produces a wrong-delimiter warning.
    /// </summary>
    [Fact]
    public void Csv_TabFileValidatedAsComma_WarnsAboutDelimiter()
    {
        var result = SpreadsheetValidator.ValidateDelimited("Name\tScore\nAlice\t95\nBob\t82\n"u8.ToArray());
        Assert.Contains("tab", AssertHasIssue(result, "CSV_POSSIBLE_WRONG_DELIMITER", ValidationSeverity.Warning).Message);
    }

    /// <summary>
    /// Invalid UTF-8 byte sequences are an error.
    /// </summary>
    [Fact]
    public void Csv_InvalidUtf8_IsError()
    {
        byte[] bytes = [.. "a,b\n1,"u8.ToArray(), 0xC3, 0x28, (byte)'\n'];
        AssertHasIssue(SpreadsheetValidator.ValidateDelimited(bytes), "CSV_ENCODING_INVALID", ValidationSeverity.Error);
    }

    /// <summary>
    /// A ZIP file validated as CSV is rejected as non-text.
    /// </summary>
    [Fact]
    public void Csv_GivenXlsx_IsNotText()
    {
        var issue = AssertHasIssue(SpreadsheetValidator.ValidateDelimited(MinimalXlsx()), "CSV_NOT_TEXT", ValidationSeverity.Error);
        Assert.Contains("XLSX", issue.Message);
    }

    /// <summary>
    /// Empty files are an error.
    /// </summary>
    [Fact]
    public void Csv_Empty_IsError()
    {
        AssertHasIssue(SpreadsheetValidator.ValidateDelimited([]), "CSV_EMPTY", ValidationSeverity.Error);
        AssertHasIssue(SpreadsheetValidator.ValidateDelimited("\r\n\r\n"u8.ToArray()), "CSV_EMPTY", ValidationSeverity.Error);
    }

    /// <summary>
    /// Blank lines between rows are a warning, but trailing blank lines are ignored.
    /// </summary>
    [Fact]
    public void Csv_BlankLines_WarnOnlyBetweenRows()
    {
        AssertNoIssues(SpreadsheetValidator.ValidateDelimited("a,b\n1,2\n\n\n"u8.ToArray()));
        AssertHasIssue(SpreadsheetValidator.ValidateDelimited("a,b\n\n1,2\n"u8.ToArray()), "CSV_BLANK_LINE", ValidationSeverity.Warning);
    }

    // --- Report formatting ---

    /// <summary>
    /// ToString produces a report naming the format, counts, codes, and locations.
    /// </summary>
    [Fact]
    public void ToString_IncludesSummaryAndIssues()
    {
        var bytes = MinimalXlsx("<sheetData><row r=\"3\"><c r=\"B3\"><v>Alice</v></c></row></sheetData>");

        var report = SpreadsheetValidator.ValidateXlsx(bytes).ToString();

        Assert.StartsWith("The file is NOT a valid XLSX file (1 error, 0 warnings):", report);
        Assert.Contains("[Error] XLSX_CELL_NOT_NUMERIC (xl/worksheets/sheet1.xml, Sheet1!B3", report);
    }
}
