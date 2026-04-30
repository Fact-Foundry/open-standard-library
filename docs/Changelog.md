# Changelog

All notable changes to Open Standard Library will be documented in this file.

---

## Unreleased

### Added
- **Column width control** — Added `SetColumnWidth(int column, double width)` and `AutoFitColumns(double minWidth, double maxWidth)` to `oSpreadsheet`. XLSX uses `<cols><col>` elements, ODS generates per-column styles on `<table:table-column>`. Auto-fit uses a character-length heuristic with configurable min/max constraints
- **Freeze panes** — Added `FreezeRows` and `FreezeColumns` properties to `oSpreadsheet` with full generate/import support for both XLSX (`<pane>` element) and ODS (`settings.xml` config items)
- **Cell Styling API** — New `CellStyle` class on `oCell.Style` with support for bold, italic, underline, font color, background color, font name/size, and borders (thin/medium/thick with color per edge). Styles are deduplicated and written as dynamic `styles.xml` entries in XLSX and named automatic styles in ODS
- **Boolean cell value type** — Added `CellValueType.Boolean` with full generate/import support for both XLSX (`t="b"`) and ODS (`office:boolean-value`). Values stored as `"true"`/`"false"` strings
- **Test project** — Added `OslSpreadsheet.Tests` with 97 xUnit tests covering workbook creation, sheet/cell operations, boolean values, cell styling, freeze panes, column widths, and round-trip generate/import for ODS, XLSX, and CSV formats

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
