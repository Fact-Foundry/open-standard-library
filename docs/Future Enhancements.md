# Future Enhancements

Items planned for future development.

---

## Security and Robustness

Findings from a code review on 2026-11-01. None involve memory leaks or unmanaged resources; they concern resource exhaustion and data integrity. Resolved the same day: ODS generation made sparse, and cell positions below 1 rejected by the model and by XLSX import, and the import size limit hardened against overflow. `ToArray()` and `AutoFitColumns()` stay dense by design.

- **ODS serialization alters data and hides failures** — `XmlService.ConvertToXmlAsync` runs `.Replace("utf-8", "UTF-8")` and `.Replace("\" />", "\"/>")` over the entire serialized document, so a cell value containing `utf-8` or `" />` is changed in the file. It also catches every exception and returns an empty array, so a serializer failure produces a zero-byte `content.xml` with no error; `OdsFileService.GenerateFileAsync` then rethrows a generic message without the inner exception. Fix: limit the replacements to the XML declaration (or drop them; both forms are valid), let exceptions surface, and pass the original as `InnerException`.

- **Streaming CSV reader has no limits** — `ReadCsvRowsAsync` is streaming by design and `ImportOptions` doesn't apply, so a single field with no delimiters grows one `StringBuilder` to the input size. Decision: leave as is; add a doc comment saying the limits don't apply to this method.

- **Minor robustness** — (a) XLSX import: a negative shared-string index (`<v>-1</v>`) throws instead of yielding an empty string the way an out-of-range positive index does; the exception is wrapped, so it's only a nuisance. (b) `oSpreadsheet._AddCell` catches every exception and rethrows a generic one, discarding the cause; remove the try/catch or pass the original as `InnerException`.
