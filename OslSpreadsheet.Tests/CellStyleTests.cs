using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

public class CellStyleTests
{
    [Fact]
    public void CellStyle_DefaultValues()
    {
        var style = new CellStyle();

        Assert.False(style.Bold);
        Assert.False(style.Italic);
        Assert.False(style.Underline);
        Assert.Null(style.FontColor);
        Assert.Null(style.BackgroundColor);
        Assert.Null(style.FontName);
        Assert.Null(style.FontSize);
        Assert.Null(style.BorderTop);
        Assert.Null(style.BorderBottom);
        Assert.Null(style.BorderLeft);
        Assert.Null(style.BorderRight);
    }

    [Fact]
    public void Cell_StyleProperty_DefaultsToNull()
    {
        var cell = new oCell(1, 1);
        Assert.Null(cell.Style);
    }

    [Fact]
    public void Cell_StyleProperty_CanBeSet()
    {
        var cell = new oCell(1, 1);
        cell.Style = new CellStyle { Bold = true, BackgroundColor = "#2196F3" };

        Assert.NotNull(cell.Style);
        Assert.True(cell.Style.Bold);
        Assert.Equal("#2196F3", cell.Style.BackgroundColor);
    }

    [Fact]
    public void CellBorder_DefaultStyle_IsThin()
    {
        var border = new CellBorder();
        Assert.Equal(BorderStyle.Thin, border.Style);
        Assert.Null(border.Color);
    }

    // --- XLSX styling ---

    [Fact]
    public async Task Xlsx_StyledCells_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Styled");

        var header = sheet.AddCell(1, 1, "Header");
        header.Style = new CellStyle
        {
            Bold = true,
            FontColor = "#FFFFFF",
            BackgroundColor = "#2196F3"
        };

        sheet.AddCell(2, 1, "Data");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        Assert.Equal("Header", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Data", workbook.Sheets[0].GetRow(2).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Xlsx_BoldItalicUnderline_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Formatted");

        var cell = sheet.AddCell(1, 1, "Formatted");
        cell.Style = new CellStyle { Bold = true, Italic = true, Underline = true };

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        Assert.Equal("Formatted", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Xlsx_FontNameAndSize_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Fonts");

        var cell = sheet.AddCell(1, 1, "Custom Font");
        cell.Style = new CellStyle { FontName = "Arial", FontSize = 14 };

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        Assert.Equal("Custom Font", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Xlsx_Borders_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Borders");

        var cell = sheet.AddCell(1, 1, "Bordered");
        cell.Style = new CellStyle
        {
            BorderBottom = new CellBorder { Style = BorderStyle.Medium, Color = "#1565C0" },
            BorderTop = new CellBorder { Style = BorderStyle.Thin }
        };

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        Assert.Equal("Bordered", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Xlsx_MixedStyledAndUnstyled_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Mixed");

        var styled = sheet.AddCell(1, 1, "Styled");
        styled.Style = new CellStyle { Bold = true, BackgroundColor = "#FF0000" };

        sheet.AddCell(1, 2, "Unstyled");

        var alsoStyled = sheet.AddCell(1, 3, "Also Styled");
        alsoStyled.Style = new CellStyle { Italic = true, FontColor = "#00FF00" };

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var row = workbook.Sheets[0].GetRow(1);
        Assert.Equal("Styled", row.First(c => c.Column == 1).Value);
        Assert.Equal("Unstyled", row.First(c => c.Column == 2).Value);
        Assert.Equal("Also Styled", row.First(c => c.Column == 3).Value);
    }

    [Fact]
    public async Task Xlsx_SameStyleReused_DeduplicatesStyles()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Dedup");

        for (int r = 1; r <= 5; r++)
        {
            var cell = sheet.AddCell(r, 1, $"Row {r}");
            cell.Style = new CellStyle { Bold = true, BackgroundColor = "#F5F5F5" };
        }

        var bytes = await spreadsheet.GenerateXlsxFileAsync();
        Assert.True(bytes.Length > 0);

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        Assert.Equal(5, workbook.Sheets[0].RowCount);
    }

    // --- ODS styling ---

    [Fact]
    public async Task Ods_StyledCells_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Styled");

        var header = sheet.AddCell(1, 1, "Header");
        header.Style = new CellStyle
        {
            Bold = true,
            FontColor = "#FFFFFF",
            BackgroundColor = "#2196F3"
        };

        sheet.AddCell(2, 1, "Data");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        Assert.Equal("Header", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Data", workbook.Sheets[0].GetRow(2).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Ods_BoldItalicUnderline_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Formatted");

        var cell = sheet.AddCell(1, 1, "Formatted");
        cell.Style = new CellStyle { Bold = true, Italic = true, Underline = true };

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        Assert.Equal("Formatted", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Ods_Borders_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Borders");

        var cell = sheet.AddCell(1, 1, "Bordered");
        cell.Style = new CellStyle
        {
            BorderBottom = new CellBorder { Style = BorderStyle.Medium, Color = "#1565C0" },
            BorderLeft = new CellBorder { Style = BorderStyle.Thick, Color = "#000000" }
        };

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        Assert.Equal("Bordered", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Ods_MixedStyledAndUnstyled_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Mixed");

        var styled = sheet.AddCell(1, 1, "Styled");
        styled.Style = new CellStyle { Bold = true, BackgroundColor = "#FF0000" };

        sheet.AddCell(1, 2, "Unstyled");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var row = workbook.Sheets[0].GetRow(1);
        Assert.Equal("Styled", row.First(c => c.Column == 1).Value);
        Assert.Equal("Unstyled", row.First(c => c.Column == 2).Value);
    }

    // --- Text Wrapping ---

    [Fact]
    public void CellStyle_WrapText_DefaultsFalse()
    {
        var style = new CellStyle();
        Assert.False(style.WrapText);
    }

    [Fact]
    public async Task Xlsx_WrapText_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Wrap");

        var cell = sheet.AddCell(1, 1, "A long description that should wrap in the cell");
        cell.Style = new CellStyle { WrapText = true };

        sheet.AddCell(1, 2, "No wrap");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var row = workbook.Sheets[0].GetRow(1);
        Assert.Equal("A long description that should wrap in the cell", row.First(c => c.Column == 1).Value);
        Assert.Equal("No wrap", row.First(c => c.Column == 2).Value);
    }

    [Fact]
    public async Task Xlsx_WrapTextWithOtherStyles_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("WrapStyled");

        var cell = sheet.AddCell(1, 1, "Bold wrapped text");
        cell.Style = new CellStyle { Bold = true, WrapText = true, BackgroundColor = "#F5F5F5" };

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        Assert.Equal("Bold wrapped text", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    [Fact]
    public async Task Ods_WrapText_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Wrap");

        var cell = sheet.AddCell(1, 1, "A long description that should wrap in the cell");
        cell.Style = new CellStyle { WrapText = true };

        sheet.AddCell(1, 2, "No wrap");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var row = workbook.Sheets[0].GetRow(1);
        Assert.Equal("A long description that should wrap in the cell", row.First(c => c.Column == 1).Value);
        Assert.Equal("No wrap", row.First(c => c.Column == 2).Value);
    }

    [Fact]
    public async Task Ods_WrapTextWithOtherStyles_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("WrapStyled");

        var cell = sheet.AddCell(1, 1, "Bold wrapped text");
        cell.Style = new CellStyle { Bold = true, WrapText = true, BackgroundColor = "#F5F5F5" };

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        Assert.Equal("Bold wrapped text", workbook.Sheets[0].GetRow(1).First(c => c.Column == 1).Value);
    }

    // --- ClosedXML replacement scenario ---

    [Fact]
    public async Task Xlsx_ClosedXmlReplacementScenario_ProducesValidFile()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data Dictionary");

        var headerStyle = new CellStyle
        {
            Bold = true,
            FontColor = "#FFFFFF",
            BackgroundColor = "#2196F3",
            BorderBottom = new CellBorder { Style = BorderStyle.Medium, Color = "#1565C0" }
        };

        var headers = new[] { "Name", "Type", "Description", "Expression" };
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = sheet.AddCell(1, c + 1, headers[c]);
            cell.Style = headerStyle;
        }

        var zebraStyle = new CellStyle { BackgroundColor = "#F5F5F5" };

        for (int r = 2; r <= 10; r++)
        {
            sheet.AddCell(r, 1, $"Field{r - 1}");
            sheet.AddCell(r, 2, "String");
            sheet.AddCell(r, 3, $"Description for field {r - 1}");
            sheet.AddCell(r, 4, $"=EXPR({r - 1})");

            if (r % 2 == 0)
                foreach (var c in sheet.GetRow(r))
                    c.Style = zebraStyle;
        }

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var imported = workbook.Sheets[0];

        Assert.Equal("Data Dictionary", imported.SheetName);
        Assert.Equal(10, imported.RowCount);
        Assert.Equal(4, imported.ColumnCount);
        Assert.Equal("Name", imported.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Field1", imported.GetRow(2).First(c => c.Column == 1).Value);
    }
}
