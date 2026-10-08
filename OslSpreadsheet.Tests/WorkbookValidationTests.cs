using OoxSpreadsheet;
using OslSpreadsheet.Models;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests for oWorkbook.Validate() and the pre-generation check that stops invalid files from being written.
/// </summary>
public class WorkbookValidationTests
{
    private static ValidationIssue AssertSingleError(ValidationResult result, string code)
    {
        Assert.False(result.IsValid, result.ToString());
        var issue = Assert.Single(result.Errors);
        Assert.Equal(code, issue.Code);
        return issue;
    }

    /// <summary>
    /// A workbook using every value type with well-formed values is valid for every format.
    /// </summary>
    [Theory]
    [InlineData(ValidationFileFormat.Xlsx)]
    [InlineData(ValidationFileFormat.Ods)]
    [InlineData(ValidationFileFormat.Delimited)]
    public void ValidWorkbook_HasNoIssues(ValidationFileFormat format)
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Text");
        sheet.AddCell(1, 2, "95.5", CellValueType.Float);
        sheet.AddCell(1, 3, "42", CellValueType.Int64);
        sheet.AddCell(1, 4, "true", CellValueType.Boolean);
        sheet.AddCell(1, 5, "2024-01-31T13:45:00", CellValueType.DateTime);
        sheet.AddCell(2, 1, "tab\there and\nnewline");

        var result = workbook.Validate(format);

        Assert.True(result.IsValid, result.ToString());
        Assert.Empty(result.Issues);
    }

    /// <summary>
    /// An empty workbook is reported, and the report wording refers to the workbook rather than a file.
    /// </summary>
    [Fact]
    public void EmptyWorkbook_IsError()
    {
        var result = new oWorkbook().Validate();

        AssertSingleError(result, "WB_NO_SHEETS");
        Assert.StartsWith("The workbook is NOT valid for XLSX output", result.ToString());
    }

    /// <summary>
    /// Values that don't match their declared type are reported with sheet, cell, and row.
    /// </summary>
    [Theory]
    [InlineData("abc", CellValueType.Float, "WB_CELL_NOT_NUMERIC")]
    [InlineData("1,234.50", CellValueType.Float, "WB_CELL_NOT_NUMERIC")]
    [InlineData("", CellValueType.Float, "WB_CELL_NOT_NUMERIC")]
    [InlineData("1.5", CellValueType.Int64, "WB_CELL_NOT_NUMERIC")]
    [InlineData("yes", CellValueType.Boolean, "WB_CELL_NOT_BOOLEAN")]
    [InlineData("31/01/2024", CellValueType.DateTime, "WB_CELL_NOT_DATETIME")]
    public void MismatchedValueType_IsError(string value, CellValueType type, string code)
    {
        var workbook = new oWorkbook();
        workbook.AddSheet("Data").AddCell(3, 2, value, type);

        var issue = AssertSingleError(workbook.Validate(), code);

        Assert.Equal("Data", issue.Sheet);
        Assert.Equal("B3", issue.Cell);
        Assert.Equal(3, issue.Row);
    }

    /// <summary>
    /// A formula cell may have an empty value (no cached result) regardless of its type.
    /// </summary>
    [Fact]
    public void FormulaCellWithEmptyValue_IsValid()
    {
        var workbook = new oWorkbook();
        var cell = workbook.AddSheet("Data").AddCell(1, 1, "", CellValueType.Float);
        cell.Formula = "=1+1";

        Assert.True(workbook.Validate().IsValid);
    }

    /// <summary>
    /// Sheet name rules follow the target format: XLSX limits names to 31 characters, ODS does not.
    /// </summary>
    [Fact]
    public void SheetNameLength_DependsOnFormat()
    {
        var workbook = new oWorkbook();
        workbook.AddSheet("A sheet name that is longer than 31 chars").AddCell(1, 1, "x");

        AssertSingleError(workbook.Validate(ValidationFileFormat.Xlsx), "WB_SHEET_NAME_INVALID");
        Assert.True(workbook.Validate(ValidationFileFormat.Ods).IsValid);
        Assert.True(workbook.Validate(ValidationFileFormat.Delimited).IsValid);
    }

    /// <summary>
    /// Reserved characters and duplicate names are errors for both XLSX and ODS.
    /// </summary>
    [Fact]
    public void InvalidAndDuplicateSheetNames_AreErrors()
    {
        var workbook = new oWorkbook();
        workbook.AddSheet("Q1/Q2").AddCell(1, 1, "x");
        workbook.AddSheet("Data").AddCell(1, 1, "x");
        workbook.AddSheet("data").AddCell(1, 1, "x");

        foreach (var format in new[] { ValidationFileFormat.Xlsx, ValidationFileFormat.Ods })
        {
            var codes = workbook.Validate(format).Errors.Select(e => e.Code).ToList();
            Assert.Equal(new[] { "WB_SHEET_NAME_INVALID", "WB_SHEET_NAME_DUPLICATE" }, codes);
        }
    }

    /// <summary>
    /// Control characters that XML can't store are errors; tab, LF, and CR are allowed.
    /// </summary>
    [Fact]
    public void ControlCharacter_IsError()
    {
        var workbook = new oWorkbook();
        workbook.AddSheet("Data").AddCell(1, 1, "bad\u0001char");

        var issue = AssertSingleError(workbook.Validate(), "WB_TEXT_CONTROL_CHAR");
        Assert.Contains("U+0001", issue.Message);
    }

    /// <summary>
    /// Cell positions outside the spreadsheet grid are errors.
    /// </summary>
    [Fact]
    public void CellPositionOutOfRange_IsError()
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet("Data");
        sheet.Cells.Add(new oCell(0, 1) { Value = "x" });

        AssertSingleError(workbook.Validate(), "WB_CELL_POSITION_INVALID");
    }

    /// <summary>
    /// A formula referring to a sheet not in the workbook is an error for XLSX and ODS.
    /// </summary>
    [Fact]
    public void FormulaReferencingMissingSheet_IsError()
    {
        var workbook = new oWorkbook();
        var cell = workbook.AddSheet("Calc").AddCell(1, 1);
        cell.Formula = "=SUM(Sales!A1:A3)";

        var issue = AssertSingleError(workbook.Validate(), "WB_FORMULA_UNKNOWN_SHEET");
        Assert.Contains("\"Sales\"", issue.Message);
        Assert.True(workbook.Validate(ValidationFileFormat.Delimited).IsValid);
    }

    /// <summary>
    /// A multi-sheet workbook is valid for delimited output but warns that only the first sheet is written.
    /// </summary>
    [Fact]
    public void MultipleSheets_DelimitedWarnsOnly()
    {
        var workbook = new oWorkbook();
        workbook.AddSheet("First").AddCell(1, 1, "x");
        workbook.AddSheet("Second").AddCell(1, 1, "y");

        var result = workbook.Validate(ValidationFileFormat.Delimited);

        Assert.True(result.IsValid);
        Assert.Equal("WB_DELIMITED_MULTIPLE_SHEETS", Assert.Single(result.Warnings).Code);
    }

    /// <summary>
    /// Generate methods refuse an invalid workbook, and the exception carries the full result.
    /// </summary>
    [Fact]
    public async Task Generate_InvalidWorkbook_ThrowsWithResult()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(2, 1, "not a number", CellValueType.Float);

        var xlsx = await Assert.ThrowsAsync<InvalidWorkbookException>(() => spreadsheet.GenerateXlsxFileAsync());
        var ods = await Assert.ThrowsAsync<InvalidWorkbookException>(() => spreadsheet.GenerateOdsFileAsync());
        var csv = await Assert.ThrowsAsync<InvalidWorkbookException>(() => spreadsheet.GenerateCsvFileAsync());

        foreach (var ex in new[] { xlsx, ods, csv })
        {
            Assert.Equal("WB_CELL_NOT_NUMERIC", Assert.Single(ex.Result.Errors).Code);
            Assert.Contains("A2", ex.Message);
        }
    }

    /// <summary>
    /// Warnings alone do not stop generation.
    /// </summary>
    [Fact]
    public async Task Generate_WarningsOnly_Succeeds()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("First").AddCell(1, 1, "x");
        spreadsheet.Workbook.AddSheet("Second").AddCell(1, 1, "y");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        Assert.NotEmpty(bytes);
    }
}
