# Validating Files

`SpreadsheetValidator` checks whether a file is a well-formed XLSX, ODS, or delimited text file *without* importing it. It is designed for pipelines where files come from an untrusted or unreliable producer, such as an LLM. In those pipelines you want to reject a malformed file before anyone opens it, and send the producer a precise description of what to fix.

```csharp
using OslSpreadsheet.Validation;

var result = SpreadsheetValidator.Validate(bytes, "report.xlsx");

if (!result.IsValid)
    Console.WriteLine(result); // Plain-text report listing every issue
```

## Choosing the Format

| Method | Validates as |
|--------|--------------|
| `Validate(bytes, fileName)` | The format implied by the extension: `.xlsx` `.xlsm` `.xltx` `.xltm`, `.ods` `.ots`, `.csv` `.tsv` `.tab` `.psv` `.txt`. If the extension isn't recognized, the format is detected from the content. |
| `Validate(bytes, ValidationFileFormat)` | The given format. Pass `ValidationFileFormat.Unknown` to detect it from the content. |
| `ValidateXlsx(bytes)` | XLSX |
| `ValidateOds(bytes)` | ODS |
| `ValidateDelimited(bytes, ColumnDelimeter)` | Delimited text with the given delimiter |
| `DetectFormat(bytes)` | Returns the detected format without validating it |

The file name decides what the file is *supposed* to be, so validating by name catches files in the wrong format. For example, a CSV saved as `report.xlsx` is reported as:

```
The file is NOT a valid XLSX file (1 error, 0 warnings):
- [Error] XLSX_NOT_ZIP: An XLSX file must be a ZIP archive, but this file is plain text (it may be CSV or another delimited format). XLSX files can't be written as plain text.
```

The validator also recognizes ODS files saved as XLSX (and the reverse), base64-encoded ZIP data, HTML, JSON, Markdown tables, PDFs, legacy `.xls` files, Excel 2003 XML, and flat ODS (`.fods`) files.

## Reading the Result

`ValidationResult`:

| Property | Description |
|----------|-------------|
| `IsValid` | `true` when there are no errors. Warnings don't make a file invalid. |
| `Format` | The format the file was validated as |
| `Issues` | All issues, in the order they were found |
| `Errors` / `Warnings` | Issues filtered by severity |
| `Truncated` | `true` if validation stopped at `MaxIssues` |
| `ToString()` | A plain-text report, ready to show to a user or send back to an LLM |

Each `ValidationIssue` is structured, so you can process it in code or serialize it to JSON:

| Property | Example | Description |
|----------|---------|-------------|
| `Severity` | `Error` | `Error` or `Warning` (see below) |
| `Code` | `XLSX_CELL_NOT_NUMERIC` | Stable identifier for the rule that failed |
| `Message` | `The cell is numeric ... but its value "Alice" is not a number. ...` | What is wrong and, where possible, how to fix it |
| `Sheet` | `Sheet1` | Sheet name, when the issue is in a specific sheet |
| `Cell` | `B3` | A1-style cell address, when the issue is in a specific cell |
| `Row` | `3` | 1-based row (for delimited files, the record number) |
| `Part` | `xl/worksheets/sheet1.xml` | ZIP entry the issue is in (XLSX/ODS only) |
| `Location` | `Sheet1!B3, line 1` | Human-readable position, including the XML or text line |

### Errors vs. Warnings

- **Error**: the file is malformed. Spreadsheet applications refuse to open it, prompt to repair it, or silently drop or misplace data. Examples: a corrupt ZIP, a missing manifest entry, a broken relationship, text stored in a numeric cell, rows out of order, a formula containing `#REF!`, or an unterminated quote in a CSV.
- **Warning**: the file deviates from the specification or looks suspicious, but opens without losing data. Examples: an empty sheet, a cached `#N/A` or `#DIV/0!` result, a reference to an undefined style, or a stray quote inside an unquoted CSV value.

Formula error values are split by meaning. `#REF!` and `#NAME?` mean the formula itself is broken, so they are errors. `#N/A`, `#DIV/0!`, `#VALUE!` and the like can be legitimate results of correct formulas on the current data, so they are warnings.

## Options

```csharp
var options = new ValidationOptions
{
    MaxIssues = 100,                          // Stop after this many issues (default 100)
    MaxUncompressedBytes = 512L * 1024 * 1024, // Reject larger XLSX/ODS packages without decompressing (zip-bomb guard)
    Delimiter = ColumnDelimeter.Tab,          // Delimited files: null (default) infers it from the extension, else comma
    Encoding = FileEncoding.UTF8,             // Delimited files: expected text encoding
    RequireConsistentColumnCount = true       // Delimited files: rows with a different field count are errors (false = warnings)
};

var result = SpreadsheetValidator.Validate(bytes, "export.tsv", options);
```

## Example: Validate LLM Output Before Delivering It

```csharp
const int maxAttempts = 3;

for (int attempt = 1; attempt <= maxAttempts; attempt++)
{
    var (fileName, bytes) = await GenerateFileWithLlmAsync(prompt);
    var result = SpreadsheetValidator.Validate(bytes, fileName);

    if (result.IsValid)
        return (fileName, bytes); // Warnings, if any, are in result.Warnings

    // Send the report back so the model can fix the specific problems
    prompt = $"""
        The file you produced is invalid. Fix these problems and generate it again:

        {result}
        """;
}

throw new InvalidOperationException("The model did not produce a valid file.");
```

If your tool protocol prefers structured data, serialize `result.Issues` (for example with `System.Text.Json`) instead of using `ToString()`.

## Validating a Workbook Before Generating

The checks above run on a finished file. `oWorkbook.Validate()` runs a matching set of checks on the in-memory model, so problems are caught before any file exists:

```csharp
var result = spreadsheet.Workbook.Validate(ValidationFileFormat.Xlsx); // the format you intend to generate

if (!result.IsValid)
    Console.WriteLine(result); // "The workbook is NOT valid for XLSX output (1 error, 0 warnings): ..."
```

`GenerateXlsxFileAsync()`, `GenerateOdsFileAsync()`, and `GenerateCsvFileAsync()` call this automatically for their format and throw `InvalidWorkbookException` if there are errors. The exception's `Result` property is the same `ValidationResult`, and its message is the plain-text report. Warnings don't stop generation.

Rules are prefixed `WB_`:

| Code | Severity | Check |
|------|----------|-------|
| `WB_NO_SHEETS` | Error | The workbook has no sheets |
| `WB_SHEET_NAME_INVALID` | Error | Blank name, or (XLSX) longer than 31 characters or containing `[ ] : * ? / \`, or (ODS) containing `[ ] * ? : / \`. Not checked for delimited output |
| `WB_SHEET_NAME_DUPLICATE` | Error | Two sheets with the same name, ignoring case |
| `WB_CELL_POSITION_INVALID` | Error | Row or column below 1, row above 1,048,576, or column above 16,384 |
| `WB_CELL_NOT_NUMERIC` | Error | A `Float` value that isn't a plain decimal, or an `Int64` value that isn't a whole number. Formula cells may be empty |
| `WB_CELL_NOT_BOOLEAN` | Error | A `Boolean` value other than `true`, `false`, `1`, or `0` |
| `WB_CELL_NOT_DATETIME` | Error | A `DateTime` value that doesn't parse as an ISO 8601 date |
| `WB_TEXT_TOO_LONG` | Error for XLSX, Warning for ODS | Text longer than 32,767 characters |
| `WB_TEXT_CONTROL_CHAR` | Error | A control character (other than tab, line feed, or carriage return) that XML can't store |
| `WB_FORMULA_UNKNOWN_SHEET` | Error | A formula refers to a sheet that isn't in the workbook |
| `WB_NUMBER_FORMAT_INVALID` | Error | A `NumberFormat` containing a control character |
| `WB_NUMBER_FORMAT_UNSUPPORTED` | Warning (ODS only) | A `NumberFormat` that can't be translated to an ODS data style; the cell falls back to the default format |
| `WB_FREEZE_INVALID` / `WB_AUTOFILTER_INVALID` | Error | Negative freeze counts, or an auto-filter range whose end is before its start |
| `WB_DELIMITED_MULTIPLE_SHEETS` | Warning | More than one sheet when generating a delimited file; only the first is written |

## What Is Checked

Validation is structural. It checks the packaging, XML, references, and cell values that spreadsheet applications rely on to open a file cleanly. It is not a full XML Schema validation of every element and attribute.

### Formula checks

Formulas are not evaluated. The validator checks:

| Code | Severity | Check |
|------|----------|-------|
| `XLSX_FORMULA_UNKNOWN_SHEET` / `ODS_FORMULA_UNKNOWN_SHEET` | Error | The formula refers to a sheet that doesn't exist, e.g. `=Sales!A1` with no "Sales" sheet. References into other workbooks are ignored. |
| `XLSX_FORMULA_BROKEN_REFERENCE` / `ODS_FORMULA_BROKEN_REFERENCE` | Error | The formula text contains `#REF!` |
| `XLSX_CELL_FORMULA_ERROR` / `ODS_CELL_FORMULA_ERROR` | Error for `#REF!` and `#NAME?`; Warning otherwise | The file stores an error as the cell's cached result |

It does not check function names, argument counts, or whether a referenced range is sensible for the formula; if your files follow a known layout, check those against your own specification. Cached-result checks only apply to files that store results. Files generated by this library store none unless you set them, so for those only the formula-text checks apply.

### XLSX

- The file is a ZIP archive with no duplicate entries, no backslash paths, and a declared uncompressed size within the limit
- Every `.xml` and `.rels` part is well-formed XML (no DTDs)
- `[Content_Types].xml` exists, every part has a content type, overrides refer to existing parts, and the workbook, worksheet, styles, and shared strings parts have the correct content types
- Every relationship has a valid, existing target, and every `r:id` refers to a declared relationship
- `_rels/.rels` points to a workbook. The workbook has at least one visible sheet with a valid, unique name (31 characters max, none of `[ ] : * ? / \`) and a unique `sheetId`
- `workbook.xml`, worksheets, and `styles.xml` contain only known elements, in the order the schema requires (Excel rejects out-of-order elements)
- Rows and cells are in ascending order, cell references are valid and match their row, and values match their type (`n`, `s`, `b`, `inlineStr`, `str`, `e`, `d`)
- Shared string indexes and cell style indexes are in range, style indexes in `cellXfs` refer to existing fonts, fills, borders, and number formats, and text is no longer than 32,767 characters
- Column ranges, merged ranges (non-overlapping), `dimension`, and `autoFilter` references are valid

### ODS

- The file is a ZIP archive whose `mimetype` entry contains the spreadsheet (or spreadsheet template) MIME type. Not being the first entry, or being compressed, is a warning.
- `META-INF/manifest.xml` exists, has a root entry, and lists every file in the archive. LibreOffice refuses files with unlisted entries.
- `content.xml`, `styles.xml`, `meta.xml`, and `settings.xml` are well-formed and have the correct root elements, and `content.xml` contains an `office:spreadsheet` body with at least one sheet
- Sheet names are unique and valid, and rows and cells are nested correctly
- Cell values match their `office:value-type` (`float`, `percentage`, `currency`, `date`, `time`, `boolean`, `string`)
- Repeat and span counts are positive integers, style references resolve, and formatting properties use the `style:` namespace

### Delimited

- The file is text in the expected encoding (invalid byte sequences, binary data, and ZIP or PDF files are rejected)
- Comma-delimited files follow RFC 4180 quoting. Quoted values must be closed, and quotes inside them must be doubled. Tab- and pipe-delimited files may optionally use the same quoting. ASCII-delimited files use the unit (0x1F) and record (0x1E) separators with no quoting.
- Every row has the same number of fields as the first row
- Warnings for blank lines between rows, stray quotes in unquoted values, values longer than 32,767 characters, and files that appear to use a different delimiter
