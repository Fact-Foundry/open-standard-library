# Working with Open Document Standard

## Create ODS File

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

    // Convert spreadsheet to ODS file
    var odsFile = await spreadsheet.GenerateOdsFileAsync();

    // Save file
    await File.WriteAllBytesAsync(@"C:\Temp\New File.ods", odsFile);
}
```

## Import ODS File

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var file = File.ReadAllBytes(@"C:\Temp\New File.ods");

    var workbook = await spreadsheet.ImportOdsFileAsync(file);

    // Code to work with the workbook
}
```

## Supported Features

ODS generation and import supports:

- Multiple sheets
- Cell value types: String, Float, Boolean
- Cell styling (bold, italic, underline, font color, background color, font name, font size, borders, text wrapping)
- Freeze panes
- Auto filters
- Column widths
- Formulas

---

[Home](README.md) | [Using the Library](Using-the-Library.md)
