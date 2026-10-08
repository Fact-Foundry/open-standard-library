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
- `AsFloat(double value)` — sets the cell's value to the number (always with a period as the decimal separator) and marks it as `Float`

## Cell Styling

Cells support styling including bold, italic, underline, font color, background color, font name, font size, text wrapping, and borders (thin/medium/thick with color per edge). Styles are applied via the `CellStyle` class on `oCell.Style`.

Styles are read back on import from both XLSX and ODS, so a file can be imported, changed, and generated again without losing its formatting. Only the properties `CellStyle` models are kept; alignment, fonts chosen by theme, and other formatting the library doesn't write are dropped. Cells whose formatting comes from a column default (common in files saved by LibreOffice) get it as their own `Style`. A cell with no formatting imports with `Style` null, and `CellStyle.IsDefault` tells you whether a style has any effect. ODS always stores a border color, so a border written without one comes back from an ODS file with `Color = "#000000"`.

## Number Formats

`CellStyle.NumberFormat` takes an Excel format code and controls how a number or date is displayed:

| Code | Displays `1234.5` / `2024-01-31 13:45` as |
|------|------|
| `#,##0.00` | `1,234.50` |
| `0` | `1235` |
| `0.0%` | `25.6%` (for the value `0.256`) |
| `"$"#,##0.00` | `$1,234.50` |
| `#,##0.00 "EUR"` | `1,234.50 EUR` |
| `mm/dd/yyyy` | `01/31/2024` |
| `d-mmm-yy` | `31-Jan-24` |
| `mmmm d, yyyy` | `January 31, 2024` |
| `dddd` | `Wednesday` |
| `h:mm AM/PM` | `1:45 PM` |
| `yyyy-mm-dd hh:mm` | `2024-01-31 13:45` |
| `@` | text as-is |

XLSX stores the code directly, so any code Excel accepts works there. ODS has no format-code syntax, so the code is translated to an ODS data style. The translation supports decimals, thousands separators, percentages, currency symbols (including `[$€-407]`-style locale tags, which become a plain symbol), text literals, and date/time patterns built from `y`, `m`, `d`, `h`, `s`, and `AM/PM`. It doesn't support multiple sections (`positive;negative`), colors, conditions, scientific notation, fractions, elapsed time (`[h]:mm`), or padding characters; `workbook.Validate(ValidationFileFormat.Ods)` reports those as `WB_NUMBER_FORMAT_UNSUPPORTED` warnings and the cell falls back to the default format.

On import, the format is read back into `NumberFormat` from both XLSX (including Excel's built-in formats, which files store only by id) and ODS. `DateTime` cells written without a format use the library defaults (`yyyy-mm-dd` or `yyyy-mm-dd hh:mm:ss`) and import with no `Style`.

## Freeze Panes

Sheets support freezing rows and columns via `FreezeRows` and `FreezeColumns` properties on `oSpreadsheet`.

## Auto Filters

Sheets support auto filters via `SetAutoFilter()` and `SetAutoFilter(int startRow, int startCol, int endRow, int endCol)` on `oSpreadsheet`.

## Column Width

Column widths can be controlled via `SetColumnWidth(int column, double width)` and `AutoFitColumns(double minWidth, double maxWidth)` on `oSpreadsheet`.

## Formulas

Set `Formula` on a cell using Excel A1 syntax, such as `=SUM(A1:A10)` or `='My Data'!B2*2`. Formulas are written to and read from XLSX and ODS files; for ODS they are translated to and from OpenFormula syntax automatically. Delimited files store values only.

The library does not calculate formulas. Excel and LibreOffice compute them when the file opens, but tools that only read stored values see an empty result unless you set the cell's `Value` (and `ValueType`) to the known result, which is then stored as the cached value.

## Validating Before Generating

`workbook.Validate(format)` checks the model for anything that would produce an invalid file: values that don't match their `ValueType`, invalid or duplicate sheet names, out-of-range positions, control characters, and formulas that refer to missing sheets. The Generate methods run the same check and throw `InvalidWorkbookException` on errors, so an invalid file is never written. See [Validating Files](Validating-Files.md#validating-a-workbook-before-generating).

## File Type Guides

- [Working with Delimited Files](Working-with-Delimited-Files.md)
- [Working with Open Document Standard](Working-with-Open-Document-Standard.md)
- [Working with Open Office XML](Working-with-Open-Office-XML.md)
