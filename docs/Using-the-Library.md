# Using the Library

This library uses a Workbook model created in the Spreadsheet service that is implemented through the interface `ISpreadsheet`.

You can create it directly with `new Spreadsheet()`, or register it with your dependency injection container yourself (the library does not include a registration extension):

```csharp
services.AddTransient<ISpreadsheet, Spreadsheet>();
```

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

The array is sized to the sheet's row and column counts. Positions with no cell contain `null`, not an empty string.

## Cell Helpers

`oCell` has extension methods for converting values in place:

- `FromEpochSeconds()`, `FromEpochMilliseconds()`, `ToEpochSeconds()`, `ToEpochMilliseconds()` — convert between Unix epoch timestamps and ISO 8601 DateTime values (treated as UTC)
- `AsFloat<T>(float value)` — sets the cell's value and marks it as `Float`. The type parameter is unused; it is kept for backward compatibility

## Cell Styling

Cells support styling including bold, italic, underline, font color, background color, font name, font size, text wrapping, and borders (thin/medium/thick with color per edge). Styles are applied via the `CellStyle` class on `oCell.Style`.

## Freeze Panes

Sheets support freezing rows and columns via `FreezeRows` and `FreezeColumns` properties on `oSpreadsheet`.

## Auto Filters

Sheets support auto filters via `SetAutoFilter()` and `SetAutoFilter(int startRow, int startCol, int endRow, int endCol)` on `oSpreadsheet`.

## Column Width

Column widths can be controlled via `SetColumnWidth(int column, double width)` and `AutoFitColumns(double minWidth, double maxWidth)` on `oSpreadsheet`.

## Formulas

Set `Formula` on a cell using Excel A1 syntax, such as `=SUM(A1:A10)` or `='My Data'!B2*2`. Formulas are written to and read from XLSX and ODS files; for ODS they are translated to and from OpenFormula syntax automatically. Delimited files store values only.

The library does not calculate formulas. Excel and LibreOffice compute them when the file opens, but tools that only read stored values see an empty result unless you set the cell's `Value` (and `ValueType`) to the known result, which is then stored as the cached value.

## File Type Guides

- [Working with Delimited Files](Working-with-Delimited-Files.md)
- [Working with Open Document Standard](Working-with-Open-Document-Standard.md)
- [Working with Open Office XML](Working-with-Open-Office-XML.md)
