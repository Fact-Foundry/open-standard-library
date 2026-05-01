# Working with Delimited Files

## Supported Delimiters

| Delimiter | File Type |
|-----------|-----------|
| Comma | CSV |
| Tab | TXT |
| Pipe | TXT |

## Create Delimited File

To convert a Workbook to a delimited file, use the following example.

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var workbook = spreadsheet.Workbook;

    // Set to Comma by default. Other options are listed in the table above.
    workbook.ColumnDelimeter = ColumnDelimeter.Comma;

    var sheet1 = await workbook.AddSheetAsync();

    sheet1.AddCell(1, 1, "Item #");
    sheet1.AddCell(1, 2, "Price");
    sheet1.AddCell(2, 1, "5\" Fitting");
    sheet1.AddCell(2, 2, "10.20", CellValueType.Float);

    // Convert spreadsheet to delimited file
    var csvFile = await spreadsheet.GenerateCsvFileAsync();

    // Save file
    await File.WriteAllBytesAsync(@"C:\Temp\New File.csv", csvFile);
}
```

When converting to a CSV file, the values will be wrapped in double-quotes and separated by commas. Additionally, double-quotes inside a column's text will be properly escaped.

Other delimited types do not escape double-quotes and do not wrap values in double-quotes.

The code above generates the following output:

```
"Item #","Price"
"5"" Fitting","10.20"
```

## Import Delimited File

To import a CSV file into a workbook, the CSV must be in a valid format. The following text is an example of properly formatted text:

```
"Item #","Price"
"5"" Fitting","10.20"
```

To import a CSV file into a workbook, use the following code:

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var file = File.ReadAllBytes(@"C:\Temp\New File.csv");

    var workbook = await spreadsheet.ImportCsvFileAsync(file);

    // Code to work with the workbook
}
```

---

[Home](README.md) | [Using the Library](Using-the-Library.md)
