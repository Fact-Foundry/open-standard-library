# Future Enhancements

Items planned for future development.

---

## Import Enhancements

- **Sheet visibility** — Add a `Visibility` property to `oSpreadsheet` (enum `SheetVisibility`: `Visible`, `Hidden`, `VeryHidden`; default `Visible`) so callers can tell which sheets are hidden. Today the XLSX importer reads each sheet's name but ignores its `state` attribute, and the model has nowhere to store it.
  - XLSX import: read `state` on each `<sheet>` element in `xl/workbook.xml` (`visible`, `hidden`, `veryHidden`; missing means visible)
  - XLSX generate: write `state` for non-visible sheets
  - ODS import and generate: map to `table:display="false"` on the table's style. ODS has no "very hidden", so treat it as hidden on write and report a warning from `Validate(ValidationFileFormat.Ods)`
  - Validation: add a `WB_` error when no sheet is visible, matching the existing file validator rule
  - Delimited export: decide whether hidden sheets are skipped when choosing the sheet to export, and document it
  - Tests for import, generate and round-trip in both formats, plus a Changelog entry

- **Formatted cell values** — Add a helper, e.g. `oCell.GetFormattedValue(CultureInfo? culture = null)`, that returns a cell's value as a spreadsheet app would display it by applying `Style.NumberFormat`: `1234.5` with `"$"#,##0.00` becomes `$1,234.50`, and `2021-03-15T00:00:00` with `mm/dd/yyyy` becomes `03/15/2021`. Import already reads the format code, but nothing applies it, so callers that want readable text (reports, LLM context, previews) have to reimplement Excel formatting.
  - Support the same subset the ODS translation supports: decimals, thousands separators, percentages, currency symbols (including `[$€-407]` locale tags), text literals, and date/time tokens (`y`, `m`, `d`, `h`, `s`, `AM/PM`)
  - Unformatted values: `DateTime` cells use the library defaults (`yyyy-mm-dd`, or `yyyy-mm-dd hh:mm:ss` when there's a time), `Boolean` cells show `TRUE`/`FALSE`, and other cells return `Value` unchanged
  - Unsupported codes (multiple sections, colors, conditions, scientific notation, fractions, elapsed time) fall back to the raw `Value` rather than throwing. Supporting the positive/negative sections later would be a natural follow-up
  - Default to the invariant culture so output doesn't depend on the machine's locale, with an optional culture for separators and month names
  - Tests covering each supported token and the fallbacks, docs in Using-the-Library.md (next to the Number Formats table), plus a Changelog entry
