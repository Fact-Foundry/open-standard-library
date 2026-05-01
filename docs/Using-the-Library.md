# Using the Library

This library uses a Workbook model created in the Spreadsheet service that is implemented through the interface `ISpreadsheet`.

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var workbook = spreadsheet.Workbook;

    // ...
}
```

Once the Workbook model has been populated, the library can convert this model to different file types. The following sections show examples of how to use each one.

## Converting to a 2D Array

Workbooks can be converted into a 2D array:

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var workbook = spreadsheet.Workbook;

    // Converts the first sheet to a 2D array
    var sheet1 = workbook.Sheets.First().ToArray();
}
```

## Cell Styling

Cells support styling including bold, italic, underline, font color, background color, font name, font size, text wrapping, and borders (thin/medium/thick with color per edge). Styles are applied via the `CellStyle` class on `oCell.Style`.

## Freeze Panes

Sheets support freezing rows and columns via `FreezeRows` and `FreezeColumns` properties on `oSpreadsheet`.

## Auto Filters

Sheets support auto filters via `SetAutoFilter()` and `SetAutoFilter(int startRow, int startCol, int endRow, int endCol)` on `oSpreadsheet`.

## Column Width

Column widths can be controlled via `SetColumnWidth(int column, double width)` and `AutoFitColumns(double minWidth, double maxWidth)` on `oSpreadsheet`.

## File Type Guides

- [Working with Delimited Files](Working-with-Delimited-Files.md)
- [Working with Open Document Standard](Working-with-Open-Document-Standard.md)
- [Working with Open Office XML](Working-with-Open-Office-XML.md)
