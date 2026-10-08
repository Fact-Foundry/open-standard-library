using System.IO.Compression;
using System.Text;
using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests that cell formulas are written to and read from XLSX and ODS files, including translation to ODS OpenFormula syntax.
/// </summary>
public class FormulaTests
{
    // --- Helpers ---

    /// <summary>
    /// Builds a two-sheet workbook ("Calc" and "My Data") with one formula in Calc!A1 and numbers in 'My Data'!A1:A3.
    /// </summary>
    private static Spreadsheet BuildFormulaSpreadsheet(string formula, string cachedValue = "", CellValueType valueType = CellValueType.String)
    {
        var spreadsheet = new Spreadsheet();
        var calc = spreadsheet.Workbook.AddSheet("Calc");
        var data = spreadsheet.Workbook.AddSheet("My Data");
        data.AddCell(1, 1, "10", CellValueType.Float);
        data.AddCell(2, 1, "20", CellValueType.Float);
        data.AddCell(3, 1, "30", CellValueType.Float);

        var cell = calc.AddCell(1, 1, cachedValue, valueType);
        cell.Formula = formula;
        return spreadsheet;
    }

    private static string ReadEntry(byte[] package, string path)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry(path)!.Open());
        return reader.ReadToEnd();
    }

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

    // --- XLSX ---

    /// <summary>
    /// The formula is written without its leading '=' and the workbook asks the application to recalculate on load.
    /// </summary>
    [Fact]
    public async Task Xlsx_WritesFormulaAndRecalculateOnLoad()
    {
        using var spreadsheet = BuildFormulaSpreadsheet("=SUM('My Data'!A1:A3)");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.Contains("<c r=\"A1\"><f>SUM(&apos;My Data&apos;!A1:A3)</f></c>", ReadEntry(bytes, "xl/worksheets/sheet1.xml"));
        Assert.Contains("<calcPr fullCalcOnLoad=\"1\"/>", ReadEntry(bytes, "xl/workbook.xml"));
    }

    /// <summary>
    /// Workbooks without formulas don't request recalculation.
    /// </summary>
    [Fact]
    public async Task Xlsx_NoFormulas_NoCalcPr()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "x");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.DoesNotContain("calcPr", ReadEntry(bytes, "xl/workbook.xml"));
    }

    /// <summary>
    /// A cell value on a formula cell is written as the cached result, typed by the cell's value type.
    /// </summary>
    [Theory]
    [InlineData(CellValueType.Float, "60", "<c r=\"A1\"><f>SUM(B1:B3)</f><v>60</v></c>")]
    [InlineData(CellValueType.String, "abc", "<c r=\"A1\" t=\"str\"><f>SUM(B1:B3)</f><v>abc</v></c>")]
    [InlineData(CellValueType.Boolean, "true", "<c r=\"A1\" t=\"b\"><f>SUM(B1:B3)</f><v>1</v></c>")]
    public async Task Xlsx_WritesCachedValue(CellValueType type, string value, string expected)
    {
        using var spreadsheet = BuildFormulaSpreadsheet("SUM(B1:B3)", value, type);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.Contains(expected, ReadEntry(bytes, "xl/worksheets/sheet1.xml"));
    }

    /// <summary>
    /// Formulas and cached values survive an XLSX round-trip, and a missing leading '=' is added.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_PreservesFormula()
    {
        using var spreadsheet = BuildFormulaSpreadsheet("SUM('My Data'!A1:A3)", "60", CellValueType.Float);
        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var cell = (await importer.ImportXlsxFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Equal("=SUM('My Data'!A1:A3)", cell.Formula);
        Assert.Equal("60", cell.Value);
    }

    /// <summary>
    /// XLSX shared formulas are expanded for each cell, shifting relative references and keeping absolute ones.
    /// </summary>
    [Fact]
    public async Task Xlsx_Import_ExpandsSharedFormulas()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "placeholder");
        var bytes = ReplaceEntry(await spreadsheet.GenerateXlsxFileAsync(), "xl/worksheets/sheet1.xml", xml =>
            System.Text.RegularExpressions.Regex.Replace(xml, "<sheetData>.*</sheetData>",
                "<sheetData>" +
                "<row r=\"1\"><c r=\"B1\"><f t=\"shared\" ref=\"B1:C2\" si=\"0\">A1*$A$1+Other!A1</f><v>1</v></c><c r=\"C1\"><f t=\"shared\" si=\"0\"/><v>2</v></c></row>" +
                "<row r=\"2\"><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>3</v></c></row>" +
                "</sheetData>"));

        using var importer = new Spreadsheet();
        var cells = (await importer.ImportXlsxFileAsync(bytes)).Sheets[0].Cells;

        Assert.Equal("=A1*$A$1+Other!A1", cells.Single(c => c.Row == 1 && c.Column == 2).Formula);
        Assert.Equal("=B1*$A$1+Other!B1", cells.Single(c => c.Row == 1 && c.Column == 3).Formula);
        Assert.Equal("=A2*$A$1+Other!A2", cells.Single(c => c.Row == 2 && c.Column == 2).Formula);
    }

    // --- ODS ---

    /// <summary>
    /// Excel A1 formulas are translated to OpenFormula: bracketed references, sheet prefixes, and ';' argument separators.
    /// Commas inside string literals are left alone.
    /// </summary>
    [Theory]
    [InlineData("=SUM(A1:A3)", "of:=SUM([.A1:.A3])")]
    [InlineData("='My Data'!$A$2*2", "of:=[$'My Data'.$A$2]*2")]
    [InlineData("=SUM('My Data'!A:A)", "of:=SUM([$'My Data'.A:.A])")]
    [InlineData("=IF(A1>50,\"big, really\",\"small\")", "of:=IF([.A1]>50;\"big, really\";\"small\")")]
    [InlineData("=LOG10(100)+A1", "of:=LOG10(100)+[.A1]")]
    [InlineData("=ROUND(AVERAGE(B2:B9),1)", "of:=ROUND(AVERAGE([.B2:.B9]);1)")]
    public async Task Ods_TranslatesFormulaToOpenFormula(string formula, string expected)
    {
        using var spreadsheet = BuildFormulaSpreadsheet(formula);

        var content = ReadEntry(await spreadsheet.GenerateOdsFileAsync(), "content.xml");

        var escaped = expected.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        Assert.Contains($"table:formula=\"{escaped}\"", content);
    }

    /// <summary>
    /// A formula cell without a cached value is written untyped so the application calculates it.
    /// </summary>
    [Fact]
    public async Task Ods_FormulaWithoutCachedValue_IsUntyped()
    {
        using var spreadsheet = BuildFormulaSpreadsheet("=1+1", "", CellValueType.Float);

        var content = ReadEntry(await spreadsheet.GenerateOdsFileAsync(), "content.xml");

        Assert.Matches("<table:table-cell table:style-name=\"ce1\" table:formula=\"of:=1\\+1\"\\s*/>", content);
    }

    /// <summary>
    /// Formulas and cached values survive an ODS round-trip, translated back to Excel A1 syntax.
    /// </summary>
    [Theory]
    [InlineData("=SUM('My Data'!A1:A3)")]
    [InlineData("='My Data'!$A$2*2")]
    [InlineData("=IF(A1>50,\"big, really\",\"small\")")]
    [InlineData("=COUNTIF('My Data'!A1:A3,\">15\")")]
    public async Task Ods_RoundTrip_PreservesFormula(string formula)
    {
        using var spreadsheet = BuildFormulaSpreadsheet(formula, "60", CellValueType.Float);
        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var cell = (await importer.ImportOdsFileAsync(bytes)).Sheets[0].Cells.Single();

        Assert.Equal(formula, cell.Formula);
        Assert.Equal("60", cell.Value);
    }

    // --- Validation ---

    /// <summary>
    /// Files with formulas generated by the library validate cleanly in both formats.
    /// </summary>
    [Fact]
    public async Task Generated_FormulaFiles_AreValid()
    {
        using var spreadsheet = BuildFormulaSpreadsheet("=SUM('My Data'!A1:A3)");

        var xlsx = SpreadsheetValidator.ValidateXlsx(await spreadsheet.GenerateXlsxFileAsync());
        var ods = SpreadsheetValidator.ValidateOds(await spreadsheet.GenerateOdsFileAsync());

        Assert.True(xlsx.Issues.Count == 0, xlsx.ToString());
        Assert.True(ods.Issues.Count == 0, ods.ToString());
    }

    /// <summary>
    /// A formula referring to a sheet that doesn't exist is an error in both formats.
    /// </summary>
    [Fact]
    public async Task Validator_FormulaReferencingMissingSheet_IsError()
    {
        // Generation refuses such a workbook, so write a valid file and rename the referenced sheet in the stored formula
        using var spreadsheet = BuildFormulaSpreadsheet("=SUM('My Data'!A1:A3)+'My Data'!A1");
        var xlsxBytes = ReplaceEntry(await spreadsheet.GenerateXlsxFileAsync(), "xl/worksheets/sheet1.xml", xml => xml.Replace("SUM(&apos;My Data&apos;!A1:A3)", "SUM(Sales!A1:A3)"));
        var odsBytes = ReplaceEntry(await spreadsheet.GenerateOdsFileAsync(), "content.xml", xml => xml.Replace("SUM([$'My Data'.A1:.A3])", "SUM([$Sales.A1:.A3])"));

        var xlsx = SpreadsheetValidator.ValidateXlsx(xlsxBytes);
        var ods = SpreadsheetValidator.ValidateOds(odsBytes);

        foreach (var (result, code) in new[] { (xlsx, "XLSX_FORMULA_UNKNOWN_SHEET"), (ods, "ODS_FORMULA_UNKNOWN_SHEET") })
        {
            var issue = Assert.Single(result.Issues);
            Assert.Equal(code, issue.Code);
            Assert.Equal(ValidationSeverity.Error, issue.Severity);
            Assert.Contains("\"Sales\"", issue.Message);
            Assert.Equal("Calc", issue.Sheet);
            Assert.Equal("A1", issue.Cell);
        }
    }

    /// <summary>
    /// References into other workbooks ([1]Sheet!A1) are not checked against this workbook's sheets.
    /// </summary>
    [Fact]
    public async Task Validator_ExternalWorkbookReference_IsNotFlagged()
    {
        using var spreadsheet = BuildFormulaSpreadsheet("=[1]Budget!A1");

        var result = SpreadsheetValidator.ValidateXlsx(await spreadsheet.GenerateXlsxFileAsync());

        Assert.DoesNotContain(result.Issues, i => i.Code == "XLSX_FORMULA_UNKNOWN_SHEET");
    }
}
