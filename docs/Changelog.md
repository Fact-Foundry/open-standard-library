# Changelog

All notable changes to Open Standard Library will be documented in this file.

---

## Unreleased

### Added
- **Formulas in XLSX and ODS** — `oCell.Formula` (Excel A1 syntax, e.g. `=SUM('My Data'!A1:A3)`) is now written when generating and read when importing. Previously the property was stored but never written to files. ODS formulas are translated to and from OpenFormula syntax (`of:=SUM([$'My Data'.A1:.A3])`). XLSX shared formulas are expanded on import. A cell's `Value` on a formula cell is stored as the cached result; without one, XLSX workbooks are flagged to recalculate on load and ODS cells are left untyped so the application calculates them. The library does not calculate formulas itself
- **Validator: formulas referencing missing sheets** — Added `XLSX_FORMULA_UNKNOWN_SHEET` and `ODS_FORMULA_UNKNOWN_SHEET` errors for formulas that refer to a sheet not in the workbook
- **Delimited import for every delimiter** — `ImportCsvFileAsync()` and `ReadCsvRowsAsync()` now use the workbook's `ColumnDelimeter`, so tab-, pipe-, and ASCII-delimited files can be imported
- **`GetCell(int row, int column)`** on `oSpreadsheet` — returns the cell at a position, or null
- **Styles and column widths on import** — XLSX and ODS import now read cell formatting (bold, italic, underline, font color/name/size, background color, borders, wrap, number format) into `oCell.Style`, and explicitly set column widths into `ColumnWidths`, so a file can be edited and regenerated without losing its formatting. Cells formatted through a column default (as LibreOffice writes) are handled. Added `CellStyle.IsDefault` and `Clone()`, and `CellBorder.Clone()`
- **Number formats** — Added `CellStyle.NumberFormat`, an Excel format code such as `#,##0.00`, `0%`, `"$"#,##0.00`, `mmm d, yyyy`, or `h:mm AM/PM`. XLSX writes the code to `styles.xml` (using Excel's built-in ids where one matches) and ODS translates it to a `number:*-style` data style, with percentage and currency cells given the matching ODS value type and `office:currency`. Both importers read formats back, including Excel's built-in formats that files store only by id. Formats ODS can't express (scientific, fractions, elapsed time, multiple sections, colors) are reported by `Validate()` as `WB_NUMBER_FORMAT_UNSUPPORTED` warnings and fall back to the default. See [Using the Library](Using-the-Library.md#number-formats)
- **Workbook validation before generating** — Added `oWorkbook.Validate(format)`, which checks the in-memory model for problems that would produce an invalid file: empty workbooks, invalid or duplicate sheet names (with XLSX's 31-character limit applied only to XLSX), values that don't match their `CellValueType`, out-of-range cell positions, control characters, over-long text, invalid freeze or auto-filter settings, and formulas that refer to missing sheets. Returns the same `ValidationResult` as the file validator with `WB_`-prefixed codes. See [Validating Files](Validating-Files.md#validating-a-workbook-before-generating)
- **`InvalidWorkbookException`** — `GenerateXlsxFileAsync()`, `GenerateOdsFileAsync()`, and `GenerateCsvFileAsync()` now validate first and throw this exception (carrying the `ValidationResult`) instead of silently writing a file that Excel or LibreOffice would reject. Previously a `Float` cell holding `"abc"` or a 40-character sheet name produced a broken file with no error
- **Tests** — Added 21 formula tests, 16 delimited import tests, 7 ODS import tests, 2 XLSX generation tests, 3 cell index tests, 19 workbook validation tests, 44 number format tests, and 10 style import tests, bringing test count from 184 to 306

### Changed
- **Linear-time cell handling** — `oSpreadsheet` now keeps a position index, so `AddCell()` and the new `GetCell(row, column)` are O(1) instead of scanning every cell. XLSX and ODS generation no longer rescan the cell list for every row or cell. Adding 200,000 cells dropped from minutes to under 50 ms, and importing a 200,000-cell file from over a minute to under a second. `Cells` is still a public list; if it is modified directly, the index is rebuilt on the next lookup

### Fixed
- **ODS import dropped rows inside header-rows and row groups** — Rows wrapped in `table:table-header-rows` (written by LibreOffice when rows are set to repeat on printed pages), `table:table-rows`, or `table:table-row-group` (outline groups) were skipped entirely. They are now imported in document order
- **ODS import shifted columns after merged cells** — `table:covered-table-cell` elements (the cells hidden by a merge) were skipped, so every cell to the right of a merged range was imported one or more columns too far left. Covered cells now occupy their column positions
- **ODS import read formatted display text instead of numbers** — Numeric cells imported the paragraph text (e.g. `1,234.50`, `$9.99`) with a Float type instead of the stored `office:value` (`1234.5`, `9.99`). Percentage and currency cells, which were imported as display strings, now import as Float with their stored value (`25%` becomes `0.25`)
- **ODS import truncated multi-line text** — Only the first paragraph of a cell was read, so text after a line break was lost. All paragraphs are now joined with line feeds, and `text:s` (repeated spaces), `text:tab`, and `text:line-break` elements are expanded to their characters, so leading spaces and tabs survive
- **XLSX with several frozen sheets opened with the sheets grouped** — Every sheet with freeze panes was written as the selected tab. Excel treats multiple selected tabs as a group, so edits made to one sheet were applied to all of them. Only the first sheet is now marked selected
- **Whitespace-only cells lost on import** — A cell containing only spaces was imported as empty from both XLSX and ODS because whitespace-only XML text was discarded on load. It is now preserved
- **XLSX leading and trailing spaces** — Text cells are now written with `xml:space="preserve"`, so values such as `"  padded  "` keep their spaces when opened in Excel
- **CSV import of unquoted values** — Import only handled files where every value was wrapped in double quotes; unquoted files such as `a,b` were silently imported as garbage. Import now follows RFC 4180, handling quoted and unquoted values, delimiters and line breaks inside quoted values, and CRLF line endings. `ReadCsvRowsAsync()` now also reads quoted values that span lines
- **Delimited import ignored the configured encoding** — `ImportCsvFileAsync()` always decoded as UTF-8 regardless of `FileEncoding`; it now uses the workbook's setting, and the imported workbook keeps its delimiter and encoding
- **Tab and pipe values containing the delimiter** — Generated tab- and pipe-delimited files wrote such values unquoted, splitting them into extra columns. Values containing the delimiter or a line break, or starting with a quote, are now quoted
- **XLSX import of error cells** — Cells with cached error values (`t="e"`, e.g. `#N/A`) were imported as Float; they are now imported as String
- **Documentation accuracy** — ODS and XLSX docs list all supported value types. The XLSX example no longer sets `Creator`, which is only written to ODS metadata. Documented that delimited export writes only the first sheet, that `ToArray()` leaves `null` for empty positions, and the `AsFloat()` helper. Added DI registration guidance and formula documentation

---

## v1.0.3 — 2026-10-06

### Added
- **File validation** — Added `SpreadsheetValidator` (namespace `OslSpreadsheet.Validation`) to check that XLSX, ODS, and delimited files are well-formed without importing them. Returns a `ValidationResult` with `IsValid` and structured `ValidationIssue`s (severity, rule code, message, sheet, cell, row, part, location), plus a plain-text report via `ToString()` suitable for sending back to an LLM. Errors (files that applications refuse, repair, or lose data from) are separated from warnings (spec deviations that open cleanly). Detects files in the wrong format, such as CSV text saved as `.xlsx`, base64-encoded ZIP data, or ODS saved as XLSX. Includes a zip-bomb size guard and an issue limit. See [Validating Files](Validating-Files.md)
- **Tests** — Added 45 validation tests, bringing test count from 139 to 184

### Changed
- **Resolved all compiler warnings** — Fixed nullable value type warnings in `AutoFilterTests`, non-nullable property warnings in `InMemoryFile` and `ODContent`, null reference warnings in `XmlService`, unused variable in `XmlService`, and async `EndOfStream` usage in `Spreadsheet` (CA2024)
- **XML documentation** — Added XML doc comments to `ISpreadsheet`, `Spreadsheet`, `InMemoryFile`, `XmlService`, and all `AutoFilterTests` methods
- **CI/CD** — Opted into Node.js 24 for GitHub Actions to resolve Node.js 20 deprecation warnings

### Fixed
- **ODS default styles in the wrong namespace** — `styles.xml` wrote `table-cell-properties`, `text-properties`, and `header-footer-properties` in the `table:` namespace instead of `style:`, so applications ignored those default formatting properties
- **ODS row count off by one** — Generated sheets declared 1,048,577 rows (including the trailing filler row) instead of the 1,048,576-row maximum

---

## v1.0.2 — 2026-05-19

### Added
- **File encoding options** — Added `FileEncoding` enum (`UTF8`, `ASCII`, `Unicode`, `UTF32`) and `FileEncoding` property on `oWorkbook`. Delimited file export and import now use the configured encoding instead of hardcoded UTF-8. Does not affect ODS or XLSX, which require UTF-8 per their specifications
- **Dynamic version string** — Generator name and version are now derived from the assembly version at runtime, which is set automatically from the git tag during CI builds. Replaces hardcoded version strings in `oWorkbook` and ODS metadata
- **Header row detection** — Added `HasHeaderRow` property, `HeaderNames` computed property, and `GetColumn(string headerName)` method to `oSpreadsheet`. When `HasHeaderRow` is true, first-row values are exposed as column names and data cells can be retrieved by header name. XLSX and ODS import auto-detect header rows when the sheet has an autoFilter starting at row 1 or frozen panes on the first row
- **Streaming CSV row reader** — Added `ReadCsvRowsAsync(Stream, bool hasHeaderRow, int? rowLimit)` to `Spreadsheet`. Reads rows one at a time via `IAsyncEnumerable<string[]>` without loading the entire file into memory. Supports optional header row consumption (stored in `CsvHeaders`) and row limit
- **DateTime and Int64 cell value types** — Added `CellValueType.DateTime` and `CellValueType.Int64`. DateTime values are stored as ISO 8601 strings and round-trip through both XLSX (OLE Automation serial dates with numFmt style) and ODS (`office:date-value` attribute). Int64 is exported as a numeric value in both formats. XLSX import detects date-formatted cells by reading styles.xml numFmtIds
- **Epoch time conversion** — Added `FromEpochSeconds()`, `FromEpochMilliseconds()`, `ToEpochSeconds()`, and `ToEpochMilliseconds()` extension methods on `oCell`. Converts between Unix epoch timestamps and ISO 8601 DateTime values. All conversions treat values as UTC
- **Date-only formatting** — DateTime cells with date-only values (no time component) now use a `yyyy-mm-dd` format in both XLSX and ODS, instead of showing `00:00:00` for the time portion

### Fixed
- **CSV import quote-escaping bug** — Doubled quotes (`""`) are now correctly unescaped to single quotes (`"`) on CSV import per RFC-4180

---

## v1.0.1 — 2026-05-01

### Added
- **Auto-filters** — Added `SetAutoFilter()` and `SetAutoFilter(int startRow, int startCol, int endRow, int endCol)` to `oSpreadsheet`. XLSX emits `<autoFilter ref="..."/>`, ODS adds `<table:database-range>` with `display-filter-buttons="true"`. Full round-trip import/generate support for both formats
- **Tests** — Added auto-filter tests, bringing test count from 102 to 113

### Fixed
- **ODS files requiring repair in LibreOffice** — Fixed multiple issues preventing ODS files from opening cleanly: incorrect namespace on `table-column-properties`, invalid empty `number-columns-repeated` attribute, UTF-8 BOM in XML files, `standalone="yes"` in XML declarations, missing `settings.xml` entry in manifest, and mimetype entry not stored uncompressed as first ZIP entry

---

## v1.0.0 — 2026-04-30

### Added
- **Column width control** — Added `SetColumnWidth(int column, double width)` and `AutoFitColumns(double minWidth, double maxWidth)` to `oSpreadsheet`. XLSX uses `<cols><col>` elements, ODS generates per-column styles on `<table:table-column>`. Auto-fit uses a character-length heuristic with configurable min/max constraints
- **Freeze panes** — Added `FreezeRows` and `FreezeColumns` properties to `oSpreadsheet` with full generate/import support for both XLSX (`<pane>` element) and ODS (`settings.xml` config items)
- **Cell Styling API** — New `CellStyle` class on `oCell.Style` with support for bold, italic, underline, font color, background color, font name/size, and borders (thin/medium/thick with color per edge). Styles are deduplicated and written as dynamic `styles.xml` entries in XLSX and named automatic styles in ODS
- **Text wrapping** — Added `WrapText` property to `CellStyle`. XLSX emits `<alignment wrapText="1"/>` in styles.xml, ODS sets `fo:wrap-option="wrap"` on table-cell-properties
- **Boolean cell value type** — Added `CellValueType.Boolean` with full generate/import support for both XLSX (`t="b"`) and ODS (`office:boolean-value`). Values stored as `"true"`/`"false"` strings
- **Test project** — Added `OslSpreadsheet.Tests` with 102 xUnit tests covering workbook creation, sheet/cell operations, boolean values, cell styling, text wrapping, freeze panes, column widths, and round-trip generate/import for ODS, XLSX, and CSV formats
- **CI/CD** — GitHub Actions workflow to publish to NuGet on version tags

---

## Previous (pre-changelog)

- ODS and XLSX file generation
- ODS and XLSX file import
- CSV/delimited file generation and import (comma, pipe, tab, ASCII)
- Multi-sheet workbook support
- String and Float cell value types
- Formula property on cells
