# Working with Open Document Standard

## Create ODS File

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var workbook = spreadsheet.Workbook;

    // Written to the ODS metadata (meta.xml); not used for XLSX or delimited files
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

ODS generation supports:

- Multiple sheets
- Cell value types: String, Float, Int64, Boolean, DateTime
- Cell styling (bold, italic, underline, font color, background color, font name, font size, borders, text wrapping)
- Freeze panes
- Auto filters
- Column widths
- Formulas

ODS import supports:

- Multiple sheets
- Cell values: String, Float, Boolean, DateTime (whole numbers are imported as Float)
- Freeze panes
- Auto filters, and header row detection from freeze panes or auto filters
- Formulas, along with any cached results stored in the file

Cell styles and column widths are not read on import.

Formulas are written in Excel syntax and translated to OpenFormula (`of:=SUM([.A1:.A10])`) in the file; on import they are translated back. The library does not calculate formulas, so generated files contain no cached results unless you set the cell's `Value`. Excel and LibreOffice calculate on open; tools that only read stored values (such as pandas or file previewers) show those cells as empty.

---

[Home](README.md) | [Using the Library](Using-the-Library.md)
