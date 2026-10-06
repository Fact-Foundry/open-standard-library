# Open Standard Library

A .NET library for reading and creating spreadsheet files in ODS, XLSX, and delimited formats (CSV, TSV, pipe-delimited). No external dependencies — just pure .NET.

[![NuGet](https://img.shields.io/nuget/v/OpenStandardLibrary)](https://www.nuget.org/packages/OpenStandardLibrary)

## Supported Formats

| Format | Generate | Import |
|--------|----------|--------|
| ODS (OpenDocument Spreadsheet) | Yes | Yes |
| XLSX (Office Open XML) | Yes | Yes |
| Delimited (comma, tab, pipe, ASCII) | Yes | Comma-delimited only, with every value wrapped in double quotes |

## Features

- **Multi-sheet workbooks** — create and import workbooks with multiple named sheets
- **Cell value types** — String, Float, Boolean, DateTime, and Int64
- **Cell styling** — bold, italic, underline, font color/name/size, background color, borders (thin/medium/thick with color per edge), and text wrapping. Applied when generating XLSX and ODS; not read on import
- **Column widths** — manual `SetColumnWidth()` or automatic `AutoFitColumns()` with min/max constraints. Applied when generating XLSX and ODS; not read on import
- **Freeze panes** — freeze rows and/or columns with `FreezeRows` and `FreezeColumns`
- **Auto-filters** — `SetAutoFilter()` for the full range or a custom range
- **Header row detection** — `HasHeaderRow`, `HeaderNames`, and `GetColumn(string)` for column-name-based access. Auto-detected on XLSX/ODS import when freeze panes or auto-filters are present
- **Streaming CSV reader** — `ReadCsvRowsAsync()` reads rows one at a time via `IAsyncEnumerable` with optional header detection and row limit (values must be wrapped in double quotes)
- **File encoding options** — UTF-8 (default), ASCII, Unicode (UTF-16), and UTF-32 for delimited files
- **Epoch conversion** — `FromEpochSeconds()`, `FromEpochMilliseconds()`, `ToEpochSeconds()`, `ToEpochMilliseconds()` extension methods on cells
- **Formulas** — set `Formula` on a cell using Excel syntax (`=SUM(A1:A10)`); written to and read from XLSX and ODS, with automatic translation to ODS OpenFormula syntax. The library does not calculate results — see [Formulas](#formulas)
- **File validation** — `SpreadsheetValidator` checks that XLSX, ODS, and delimited files are well-formed without importing them, returning structured errors and warnings (sheet, cell, row, rule code, message). Useful for vetting files produced by LLMs or other tools before they reach users
- **Dependency injection** — implements `ISpreadsheet` with `IDisposable`/`IAsyncDisposable`

## Installation

```
dotnet add package OpenStandardLibrary
```

## Quick Start

### Generate an XLSX file

```csharp
using OoxSpreadsheet;
using OslSpreadsheet.Models;

using var spreadsheet = new Spreadsheet();
var sheet = spreadsheet.Workbook.AddSheet("Data");

sheet.AddCell(1, 1, "Name");
sheet.AddCell(1, 2, "Score");
sheet.AddCell(2, 1, "Alice");
sheet.AddCell(2, 2, "95.5", CellValueType.Float);

var bytes = await spreadsheet.GenerateXlsxFileAsync();
await File.WriteAllBytesAsync("output.xlsx", bytes);
```

### Import and read by column name

```csharp
using var spreadsheet = new Spreadsheet();
var workbook = await spreadsheet.ImportXlsxFileAsync(File.ReadAllBytes("input.xlsx"));
var sheet = workbook.Sheets[0];

// HasHeaderRow is auto-detected from freeze panes or auto-filters
if (sheet.HasHeaderRow)
{
    var names = sheet.GetColumn("Name");
    foreach (var cell in names)
        Console.WriteLine(cell.Value);
}
```

### Validate a file

```csharp
using OslSpreadsheet.Validation;

var result = SpreadsheetValidator.Validate(File.ReadAllBytes("report.xlsx"), "report.xlsx");

if (!result.IsValid)
{
    foreach (var issue in result.Errors)
        Console.WriteLine($"{issue.Code} {issue.Sheet}!{issue.Cell}: {issue.Message}");
}
```

See [Validating Files](https://github.com/kupokev/open-standard-library/blob/main/docs/Validating-Files.md) for details.

### Stream CSV rows

```csharp
using var spreadsheet = new Spreadsheet();
using var stream = File.OpenRead("large-file.csv");

await foreach (var row in spreadsheet.ReadCsvRowsAsync(stream, hasHeaderRow: true, rowLimit: 100))
{
    Console.WriteLine(string.Join(", ", row));
}

// Headers are available after iteration
Console.WriteLine($"Columns: {string.Join(", ", spreadsheet.CsvHeaders!)}");
```

### Cell styling

```csharp
var cell = sheet.AddCell(1, 1, "Header");
cell.Style = new CellStyle
{
    Bold = true,
    FontColor = "#FFFFFF",
    BackgroundColor = "#2196F3",
    BorderBottom = new CellBorder { Style = BorderStyle.Medium, Color = "#1565C0" },
    WrapText = true
};
```

### Formulas

```csharp
sheet.AddCell(1, 1, "10", CellValueType.Float);
sheet.AddCell(2, 1, "20", CellValueType.Float);

var total = sheet.AddCell(3, 1);
total.Formula = "=SUM(A1:A2)";          // Excel A1 syntax; other sheets as 'Sheet Name'!A1
```

The library doesn't calculate formulas. Excel and LibreOffice compute them when the file is opened (generated XLSX files are flagged to recalculate on load), but tools that read stored values without calculating — such as pandas, file previewers, or this library's own import — see an empty result. If you already know the result, set it as the cell's `Value` and `ValueType`, and it will be stored as the cached result.

### Epoch conversion

```csharp
var cell = sheet.AddCell(1, 1, "1747650600");
cell.FromEpochSeconds(); // Value becomes "2025-05-19T10:30:00", ValueType becomes DateTime

// Or convert back
cell.ToEpochSeconds(); // Value becomes "1747650600", ValueType becomes Int64
```

## Supported Frameworks

- .NET 8.0
- .NET 9.0
- .NET 10.0

## Disclaimer

This software is provided "as is", without warranty of any kind. Use at your own risk. See the [LICENSE](LICENSE) file for full terms.

## License

MIT
