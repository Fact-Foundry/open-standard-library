# Changelog

All notable changes to Open Standard Library will be documented in this file.

---

## Unreleased

### Added
- **Formulas in XLSX and ODS** — `oCell.Formula` (Excel A1 syntax, e.g. `=SUM('My Data'!A1:A3)`) is now written when generating and read when importing. Previously the property was stored but never written to files. ODS formulas are translated to and from OpenFormula syntax (`of:=SUM([$'My Data'.A1:.A3])`). XLSX shared formulas are expanded on import. A cell's `Value` on a formula cell is stored as the cached result; without one, XLSX workbooks are flagged to recalculate on load and ODS cells are left untyped so the application calculates them. The library does not calculate formulas itself
- **Validator: formulas referencing missing sheets** — Added `XLSX_FORMULA_UNKNOWN_SHEET` and `ODS_FORMULA_UNKNOWN_SHEET` errors for formulas that refer to a sheet not in the workbook
- **Delimited import for every delimiter** — `ImportCsvFileAsync()` and `ReadCsvRowsAsync()` now use the workbook's `ColumnDelimeter`, so tab-, pipe-, and ASCII-delimited files can be imported
- **Tests** — Added 21 formula tests and 16 delimited import tests, bringing test count from 184 to 221

### Fixed
- **CSV import of unquoted values** — Import only handled files where every value was wrapped in double quotes; unquoted files such as `a,b` were silently imported as garbage. Import now follows RFC 4180, handling quoted and unquoted values, delimiters and line breaks inside quoted values, and CRLF line endings. `ReadCsvRowsAsync()` now also reads quoted values that span lines
- **Delimited import ignored the configured encoding** — `ImportCsvFileAsync()` always decoded as UTF-8 regardless of `FileEncoding`; it now uses the workbook's setting, and the imported workbook keeps its delimiter and encoding
- **Tab and pipe values containing the delimiter** — Generated tab- and pipe-delimited files wrote such values unquoted, splitting them into extra columns. Values containing the delimiter or a line break, or starting with a quote, are now quoted
- **XLSX import of error cells** — Cells with cached error values (`t="e"`, e.g. `#N/A`) were imported as Float; they are now imported as String
- **Documentation accuracy** — ODS and XLSX docs no longer list cell styles and column widths as imported, and list all supported value types. Added DI registration guidance and formula documentation

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
