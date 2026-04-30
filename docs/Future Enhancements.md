# Future Enhancements

Items planned for future development.

---

## Auto-Filters

**Status:** Not implemented.

**Needed features:**
- **Set auto-filter on a range** — Add dropdown filter arrows to header rows so users can filter/sort data in the exported spreadsheet

**Implementation notes:**
- Add a method like `sheet.SetAutoFilter(int startRow, int startCol, int endRow, int endCol)` or `sheet.SetAutoFilter()` for the used range
- XLSX: Add `<autoFilter ref="A1:G100"/>` to the worksheet XML
- ODS: Add `<table:database-range>` with `display-filter-buttons="true"` to the content XML

**Why:** Auto-filters are expected in any tabular data export. Without them, users have to manually add filters after opening the file.

---

## Use Case: Replacing ClosedXML in Semantic Modeler

The `ExportService` in FactFoundry's Semantic Modeler currently uses ClosedXML to generate data dictionary Excel exports. To swap ClosedXML for OpenStandardLibrary, the remaining feature needed is:

1. **Auto-filters** — Filter arrows on header rows

All other blockers (cell styling, freeze panes, column widths, text wrapping, borders, boolean values) are implemented — see `Changelog.md`.
