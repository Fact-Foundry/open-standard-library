# Working with Open Office XML

## Create XLSX File

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var workbook = spreadsheet.Workbook;

    // Create worksheets
    var sheet1 = await workbook.AddSheetAsync();
    var sheet2 = await workbook.AddSheetAsync("Stuff");

    // Add a cell to worksheet 2
    var cell = sheet2.AddCell(1, 1);
    cell.Value = "300.20";
    cell.ValueType = CellValueType.Float;

    // Convert spreadsheet to XLSX file
    var xlsxFile = await spreadsheet.GenerateXlsxFileAsync();

    // Save file
    await File.WriteAllBytesAsync(@"C:\Temp\New File.xlsx", xlsxFile);
}
```

## Import XLSX File

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var file = File.ReadAllBytes(@"C:\Temp\New File.xlsx");

    var workbook = await spreadsheet.ImportXlsxFileAsync(file);

    // Code to work with the workbook
}
```

## Supported Features

Document properties (`Creator`, `InitialCreator`, `CreationDate`) are not written to XLSX files; they are only used for ODS metadata.

XLSX generation supports:

- Multiple sheets
- Cell value types: String, Float, Int64, Boolean, DateTime
- Cell styling (bold, italic, underline, font color, background color, font name, font size, borders, text wrapping)
- Number formats (see [Using the Library](Using-the-Library.md#number-formats))
- Freeze panes
- Auto filters
- Column widths
- Formulas

XLSX import supports:

- Multiple sheets
- Cell values: String, Float, Boolean, DateTime (whole numbers are imported as Float)
- Freeze panes
- Auto filters, and header row detection from freeze panes or auto filters
- Formulas, along with any cached results stored in the file
- Number formats

Other cell styles and column widths are not read on import.

Formulas are written in Excel syntax (`=SUM(A1:A10)`), and the workbook is flagged to recalculate when opened. The library does not calculate formulas, so generated files contain no cached results unless you set the cell's `Value`. Excel and LibreOffice calculate on open; tools that only read stored values (such as pandas or file previewers) show those cells as empty.

---

[Home](README.md) | [Using the Library](Using-the-Library.md)
