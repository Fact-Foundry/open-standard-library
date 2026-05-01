# Working with Open Office XML

## Create XLSX File

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var workbook = spreadsheet.Workbook;

    workbook.Creator = "Kevin Williams";

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

XLSX generation and import supports:

- Multiple sheets
- Cell value types: String, Float, Boolean
- Cell styling (bold, italic, underline, font color, background color, font name, font size, borders, text wrapping)
- Freeze panes
- Auto filters
- Column widths
- Formulas

---

[Home](README.md) | [Using the Library](Using-the-Library.md)
