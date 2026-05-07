# Changelog

All notable changes to Open Standard Library will be documented in this file.

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

### Known Issues
- **CSV import quote-escaping bug** — Embedded double-quotes are not unescaped on import (e.g., `5"" Fitting` stays as-is instead of becoming `5" Fitting`)

---

## Previous (pre-changelog)

- ODS and XLSX file generation
- ODS and XLSX file import
- CSV/delimited file generation and import (comma, pipe, tab, ASCII)
- Multi-sheet workbook support
- String and Float cell value types
- Formula property on cells
