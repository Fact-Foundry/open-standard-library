# Working with Delimited Files

## Supported Delimiters

| Delimiter | File Type |
|-----------|-----------|
| Comma | CSV |
| Tab | TXT |
| Pipe | TXT |
| ASCII (Unit/Record Separator) | TXT |

## Create Delimited File

To convert a Workbook to a delimited file, use the following example. Delimited files have no concept of sheets, so only the first sheet in the workbook is written.

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

When converting to a CSV file, every value is wrapped in double quotes and separated by commas, and double quotes inside a value are escaped by doubling them (RFC 4180).

Tab- and pipe-delimited files write values as-is, except that a value containing the delimiter or a line break, or starting with a double quote, is wrapped in double quotes (with inner quotes doubled) so it reads back correctly. ASCII-delimited files never quote values.

The code above generates the following output:

```
"Item #","Price"
"5"" Fitting","10.20"
```

## Formula Injection

Spreadsheet applications evaluate a CSV value that starts with `=`, `+`, `-`, `@`, a tab, or a carriage return as a formula when the file is opened. A value such as `=HYPERLINK("http://evil")` or `=cmd|'/c calc'!A0` supplied by a user or an LLM therefore becomes live when the recipient opens the file ("CSV injection"). The library writes values verbatim by default. To protect recipients, turn on escaping:

```csharp
workbook.EscapeFormulaLikeValues = true;
```

Formula-like **text** values are then prefixed with an apostrophe (`'=1+1`), which Excel and LibreOffice treat as "store as text" and hide. Values that are plain numbers (`-5`, `+3.5`) and cells typed `Float`/`Int64` are left alone. The option is off by default because it changes the exported data; when it is off, `workbook.Validate(ValidationFileFormat.Delimited)` reports each such value as a `WB_DELIMITED_FORMULA_LIKE_VALUE` warning, and `SpreadsheetValidator.ValidateDelimited()` reports them in existing files as `CSV_FORMULA_LIKE_VALUE`.

Note that this library's own import keeps the apostrophe as part of the value; only spreadsheet applications strip it.

## Import Delimited File

Import uses the workbook's `ColumnDelimeter` and `FileEncoding`, so set them before importing anything other than a UTF-8 comma-delimited file:

```csharp
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    spreadsheet.Workbook.ColumnDelimeter = ColumnDelimeter.Tab; // Comma by default

    var file = File.ReadAllBytes(@"C:\Temp\New File.txt");

    var workbook = await spreadsheet.ImportCsvFileAsync(file);

    // Code to work with the workbook
}
```

Comma-, tab-, and pipe-delimited files may mix quoted and unquoted values. Quoted values can contain the delimiter, line breaks, and doubled quotes (`""`), following RFC 4180:

```
Item #,Price,Notes
"5"" Fitting",10.20,"Fits 1/2"" and
3/4"" pipe"
Elbow,4.50,
```

Import is lenient, like spreadsheet applications: a quote inside an unquoted value is kept as text, and blank lines are skipped. To detect malformed files instead, use [`SpreadsheetValidator.ValidateDelimited()`](Validating-Files.md). All values are imported as strings.

`ReadCsvRowsAsync()` uses the same rules, so it also reads quoted values that span multiple lines.

---

[Home](README.md) | [Using the Library](Using-the-Library.md)
