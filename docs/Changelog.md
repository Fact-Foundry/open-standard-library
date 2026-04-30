# Changelog

All notable changes to Open Standard Library will be documented in this file.

---

## Unreleased

### Added
- **Boolean cell value type** — Added `CellValueType.Boolean` with full generate/import support for both XLSX (`t="b"`) and ODS (`office:boolean-value`). Values stored as `"true"`/`"false"` strings
- **Test project** — Added `OslSpreadsheet.Tests` with 54 xUnit tests covering workbook creation, sheet/cell operations, boolean values, and round-trip generate/import for ODS, XLSX, and CSV formats

### Discovered
- **CSV import quote-escaping bug** — Embedded double-quotes are not unescaped on import (e.g., `5"" Fitting` stays as-is instead of becoming `5" Fitting`)

---

## Previous (pre-changelog)

- ODS and XLSX file generation
- ODS and XLSX file import
- CSV/delimited file generation and import (comma, pipe, tab, ASCII)
- Multi-sheet workbook support
- String and Float cell value types
- Formula property on cells
