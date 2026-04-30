# Future Enhancements

Items planned for future development.

---

## Cell Styling API

**Status:** Not implemented. The library has internal styling models for ODS/XLSX generation but no public API to customize cell appearance.

**Needed features:**
- **Bold / Italic / Underline** — Set font weight, style, and decoration per cell
- **Font color** — Set text color per cell (e.g., white text on dark header backgrounds)
- **Background / fill color** — Set cell background color (e.g., header rows, zebra striping, conditional highlighting)
- **Borders** — Set border style (thin, medium, thick, none), border color, and which edges (top, bottom, left, right) per cell or range
- **Font name and size** — Override the default Calibri 11pt per cell

**Implementation notes:**
- Add styling properties to `oCell` (e.g., `Bold`, `FontColor`, `BackgroundColor`, `BorderStyle`)
- Or introduce a separate `CellStyle` class that `oCell` references
- XLSX generation (`XlsxFileService.BuildStyles()`) currently creates a single hardcoded style. Needs to build a dynamic styles table from unique cell style combinations and assign style indices (`s` attribute) to each cell
- ODS generation (`ODStyles.cs`) similarly needs to map cell styles to named styles in `styles.xml`

**Why:** Required to replace ClosedXML in applications that need formatted exports (e.g., header rows with colored backgrounds, bold text, bordered data regions). Without cell styling, exports are plain data grids with no visual structure.

---

## Freeze Panes

**Status:** Not implemented.

**Needed features:**
- **Freeze rows** — Lock the top N rows so they remain visible while scrolling (e.g., freeze header row)
- **Freeze columns** — Lock the left N columns while scrolling

**Implementation notes:**
- Add `FreezeRows` and `FreezeColumns` properties to `oSpreadsheet`
- XLSX: Add `<sheetViews><sheetView><pane>` element to each worksheet XML with `ySplit`/`xSplit` attributes and `state="frozen"`
- ODS: Use `<config:config-item-map-named>` entries for horizontal/vertical split

**Why:** Header rows in exported spreadsheets become invisible as soon as the user scrolls past them. Freezing is a basic usability feature for any multi-row export.

---

## Column Width Control

**Status:** Partial infrastructure exists (ODS has a default column width property) but no public API.

**Needed features:**
- **Set column width** — Specify width for individual columns
- **Auto-fit column width** — Calculate optimal width based on cell content (requires measuring string lengths and accounting for font size)

**Implementation notes:**
- Add a method like `sheet.SetColumnWidth(int column, double width)` or a column properties model
- XLSX: Add `<cols><col min="N" max="N" width="W" customWidth="1"/>` to worksheet XML
- ODS: Set width on `<table:table-column>` style properties
- Auto-fit is harder — requires either approximate character-width calculations or accepting a "good enough" heuristic based on string length

**Why:** Without column width control, exported spreadsheets have uniform column widths that are too narrow for long values or too wide for short ones. Users have to manually resize every column.

---

## Text Wrapping

**Status:** Not implemented.

**Needed features:**
- **Wrap text** — Enable text wrapping per cell so long values display on multiple lines instead of overflowing

**Implementation notes:**
- Add a `WrapText` property to the cell style model
- XLSX: Add `<alignment wrapText="1"/>` to the cell's style definition in `styles.xml`
- ODS: Set `fo:wrap-option="wrap"` in the cell's paragraph properties

**Why:** Long text values (descriptions, expressions, formulas) are truncated or overflow in spreadsheets without wrapping. Essential for documentation-style exports.

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

## Boolean Cell Value Type

**Status:** Not directly supported. Booleans must be stored as strings.

**Needed features:**
- Add `CellValueType.Boolean` to the `CellValueType` enum
- Store as native boolean in XLSX (`<v>1</v>` / `<v>0</v>` with `t="b"`)

**Implementation notes:**
- Minor addition to `CellValueType` enum and cell serialization logic in `XlsxFileService` and `OdsFileService`

---

## Use Case: Replacing ClosedXML in Semantic Modeler

The `ExportService` in FactFoundry's Semantic Modeler currently uses ClosedXML to generate data dictionary Excel exports. To swap ClosedXML for OpenStandardLibrary, the following features are needed (in priority order):

1. **Cell styling** — Bold headers, colored header backgrounds (#2196F3 blue with white text), zebra striping (#F5F5F5 alternating rows)
2. **Freeze panes** — Freeze the header row (`FreezeRows(1)`)
3. **Column width** — Auto-fit columns with a max width cap (60 chars), minimum width (10 chars)
4. **Text wrapping** — Wrap long expression columns
5. **Auto-filters** — Filter arrows on header rows
6. **Borders** — Medium bottom border on header row (#1565C0 blue)

Features 1-3 are blockers for the swap. Features 4-6 are nice-to-haves that could be added incrementally.
