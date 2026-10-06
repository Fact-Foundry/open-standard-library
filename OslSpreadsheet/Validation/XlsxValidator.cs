using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Structural validator for Office Open XML spreadsheets (ECMA-376). Checks the problems that make Excel refuse a file
    /// or prompt to repair it: packaging, content types, relationships, schema element order, and cell values.
    /// </summary>
    internal sealed class XlsxValidator
    {
        private const string Prefix = "XLSX";
        internal const int MaxRows = 1_048_576;
        internal const int MaxColumns = 16_384;
        internal const int MaxCellTextLength = 32_767;
        private const int MaxSheetNameLength = 31;

        private static readonly XNamespace TransitionalNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace StrictNs = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        private static readonly XNamespace PackageRelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly HashSet<string> OfficeRelsNamespaces = new(StringComparer.Ordinal)
        {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships"
        };

        private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        private static readonly HashSet<string> WorkbookContentTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "application/vnd.ms-excel.addin.macroEnabled.main+xml"
        };

        // Child element order required by the ECMA-376 schema. Excel rejects out-of-order or unknown elements.
        private static readonly string[] WorkbookSequence =
        [
            "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
            "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
            "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"
        ];

        private static readonly string[] WorksheetSequence =
        [
            "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
            "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
            "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
            "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
            "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
            "extLst"
        ];

        private static readonly string[] StyleSheetSequence =
        [
            "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"
        ];

        private static readonly string[] CellSequence = ["f", "v", "is", "extLst"];

        private static readonly HashSet<string> RepeatableWorkbookElements = new() { "fileRecoveryPr" };
        private static readonly HashSet<string> RepeatableWorksheetElements = new() { "cols", "conditionalFormatting" };

        private static readonly HashSet<string> ErrorValues = new(StringComparer.Ordinal)
        {
            "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"
        };

        private readonly ZipPackage _package;
        private readonly IssueCollector _issues;
        private readonly Dictionary<string, Dictionary<string, Relationship>> _relationships = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _defaultContentTypes = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _overrideContentTypes = new(StringComparer.OrdinalIgnoreCase);
        private XNamespace _ns = TransitionalNs;
        private readonly HashSet<string> _sheetNames = new(StringComparer.OrdinalIgnoreCase);

        private XlsxValidator(ZipPackage package, IssueCollector issues)
        {
            _package = package;
            _issues = issues;
        }

        private sealed record Relationship(string Id, string Type, string Target, bool External, string? ResolvedPath, XElement Element)
        {
            /// <summary>
            /// The last segment of the relationship type URI, e.g. "worksheet". Matches both transitional and strict URIs.
            /// </summary>
            internal string Kind => Type[(Type.LastIndexOf('/') + 1)..];
        }

        internal static void Validate(byte[] file, IssueCollector issues, ValidationOptions options)
        {
            if (!FileSniffer.IsZip(file))
            {
                issues.Error($"{Prefix}_NOT_ZIP",
                    $"An XLSX file must be a ZIP archive, but this file is {FileSniffer.Describe(file)}. XLSX files can't be written as plain text.");
                return;
            }

            using var package = ZipPackage.Open(file, issues, options, Prefix);
            if (package == null)
                return;

            new XlsxValidator(package, issues).Run();
        }

        private void Run()
        {
            if (_package.Exists("mimetype") && _package.Exists("content.xml"))
            {
                _issues.Error($"{Prefix}_IS_ODS", "This file is an OpenDocument spreadsheet (ODS), not XLSX. Validate it as ODS or save it with an .ods extension.");
                return;
            }

            foreach (var part in _package.PartNames.Where(IsXmlPart))
                _package.LoadXml(part);

            if (!ValidateContentTypes())
                return;

            foreach (var relsPart in _package.PartNames.Where(n => n.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
                LoadRelationships(relsPart);

            foreach (var part in _package.PartNames.Where(IsXmlPart))
                ValidateRelationshipIdReferences(part);

            var rootRels = GetRelationships("");
            var officeDoc = rootRels?.Values.FirstOrDefault(r => r.Kind == "officeDocument");
            if (rootRels == null)
            {
                _issues.Error($"{Prefix}_ROOT_RELS_MISSING",
                    "The package is missing _rels/.rels, which must contain an officeDocument relationship pointing to the workbook (normally xl/workbook.xml).");
                return;
            }
            if (officeDoc == null)
            {
                _issues.Error($"{Prefix}_NO_OFFICE_DOCUMENT",
                    "_rels/.rels has no officeDocument relationship, so applications can't find the workbook. Add one with Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" and Target=\"xl/workbook.xml\".",
                    "_rels/.rels");
                return;
            }
            if (officeDoc.ResolvedPath == null || !_package.Exists(officeDoc.ResolvedPath))
                return; // Already reported as a missing relationship target.

            ValidateWorkbook(officeDoc.ResolvedPath);
        }

        private static bool IsXmlPart(string name) =>
            name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);

        // ---------------------------------------------------------------- Content types

        private bool ValidateContentTypes()
        {
            const string part = "[Content_Types].xml";
            if (!_package.Exists(part))
            {
                var nearMiss = _package.FindCaseInsensitive(part);
                _issues.Error($"{Prefix}_CONTENT_TYPES_MISSING", nearMiss != null
                    ? $"The package has \"{nearMiss}\" but the name must be exactly \"[Content_Types].xml\" (case-sensitive)."
                    : "The package is missing [Content_Types].xml at the root of the archive. Every XLSX file requires it.");
                return false;
            }

            var doc = _package.LoadXml(part);
            if (doc?.Root == null)
                return false;

            if (doc.Root.Name != ContentTypesNs + "Types")
            {
                _issues.Error($"{Prefix}_CONTENT_TYPES_ROOT",
                    $"The root element must be <Types xmlns=\"{ContentTypesNs.NamespaceName}\"> but is <{doc.Root.Name.LocalName}> in namespace \"{doc.Root.Name.NamespaceName}\".",
                    part);
                return false;
            }

            foreach (var el in doc.Root.Elements(ContentTypesNs + "Default"))
            {
                var ext = (string?)el.Attribute("Extension");
                var type = (string?)el.Attribute("ContentType");
                if (string.IsNullOrEmpty(ext) || string.IsNullOrEmpty(type))
                {
                    _issues.Error($"{Prefix}_CONTENT_TYPE_INVALID", "<Default> requires both Extension and ContentType attributes.", part, ZipPackage.LineOf(el));
                    continue;
                }
                if (!_defaultContentTypes.TryAdd(ext.TrimStart('.'), type))
                    _issues.Error($"{Prefix}_CONTENT_TYPE_DUPLICATE", $"Extension \"{ext}\" has more than one <Default> entry.", part, ZipPackage.LineOf(el));
            }

            foreach (var el in doc.Root.Elements(ContentTypesNs + "Override"))
            {
                var partName = (string?)el.Attribute("PartName");
                var type = (string?)el.Attribute("ContentType");
                if (string.IsNullOrEmpty(partName) || string.IsNullOrEmpty(type))
                {
                    _issues.Error($"{Prefix}_CONTENT_TYPE_INVALID", "<Override> requires both PartName and ContentType attributes.", part, ZipPackage.LineOf(el));
                    continue;
                }
                if (!partName.StartsWith('/'))
                {
                    _issues.Error($"{Prefix}_CONTENT_TYPE_INVALID", $"Override PartName \"{partName}\" must start with '/', e.g. \"/{partName}\".", part, ZipPackage.LineOf(el));
                    continue;
                }
                if (!_overrideContentTypes.TryAdd(partName, type))
                {
                    _issues.Error($"{Prefix}_CONTENT_TYPE_DUPLICATE", $"Part \"{partName}\" has more than one <Override> entry.", part, ZipPackage.LineOf(el));
                    continue;
                }
                if (_package.FindCaseInsensitive(partName[1..]) == null)
                    _issues.Error($"{Prefix}_CONTENT_TYPE_PART_MISSING",
                        $"An <Override> declares a content type for \"{partName}\", but no such part exists in the archive. Remove the override or add the part.",
                        part, ZipPackage.LineOf(el));
            }

            foreach (var name in _package.PartNames)
            {
                if (name == part)
                    continue;
                if (GetContentType(name) == null)
                    _issues.Error($"{Prefix}_CONTENT_TYPE_UNDECLARED",
                        $"Part \"{name}\" has no content type. Add an <Override PartName=\"/{name}\"> or a <Default Extension=\"{Path.GetExtension(name).TrimStart('.')}\"> to [Content_Types].xml.",
                        part);
            }

            return true;
        }

        private string? GetContentType(string partName)
        {
            if (_overrideContentTypes.TryGetValue("/" + partName, out var type))
                return type;
            var ext = Path.GetExtension(partName).TrimStart('.');
            return _defaultContentTypes.TryGetValue(ext, out type) ? type : null;
        }

        private void ExpectContentType(string partName, string expected, string description)
        {
            var actual = GetContentType(partName);
            if (actual != null && !string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
                _issues.Error($"{Prefix}_CONTENT_TYPE_WRONG",
                    $"The {description} part has content type \"{actual}\" but must be \"{expected}\".", "[Content_Types].xml", $"part /{partName}");
        }

        // ---------------------------------------------------------------- Relationships

        private static string RelsPathFor(string sourcePart)
        {
            if (sourcePart.Length == 0)
                return "_rels/.rels";
            var slash = sourcePart.LastIndexOf('/');
            var dir = slash < 0 ? "" : sourcePart[..(slash + 1)];
            var file = sourcePart[(slash + 1)..];
            return $"{dir}_rels/{file}.rels";
        }

        private static string? SourcePartFor(string relsPart)
        {
            if (relsPart == "_rels/.rels")
                return "";
            var marker = relsPart.LastIndexOf("_rels/", StringComparison.Ordinal);
            if (marker < 0 || !relsPart.EndsWith(".rels", StringComparison.Ordinal))
                return null;
            return relsPart[..marker] + relsPart[(marker + "_rels/".Length)..^".rels".Length];
        }

        private static string? ResolveTarget(string sourcePart, string target)
        {
            string combined;
            target = Uri.UnescapeDataString(target.Split('#')[0]);
            if (target.StartsWith('/'))
                combined = target[1..];
            else
            {
                var slash = sourcePart.LastIndexOf('/');
                combined = (slash < 0 ? "" : sourcePart[..(slash + 1)]) + target;
            }

            var segments = new List<string>();
            foreach (var segment in combined.Split('/'))
            {
                if (segment.Length == 0 || segment == ".")
                    continue;
                if (segment == "..")
                {
                    if (segments.Count == 0)
                        return null;
                    segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                segments.Add(segment);
            }
            return string.Join('/', segments);
        }

        private void LoadRelationships(string relsPart)
        {
            var source = SourcePartFor(relsPart);
            if (source == null)
            {
                _issues.Warning($"{Prefix}_RELS_MISPLACED", "Relationship parts must be stored in a _rels folder next to their source part.", relsPart);
                return;
            }
            if (source.Length > 0 && !_package.Exists(source))
                _issues.Warning($"{Prefix}_RELS_ORPHANED", $"This relationships part belongs to \"{source}\", which does not exist.", relsPart);

            var doc = _package.LoadXml(relsPart);
            if (doc?.Root == null)
                return;

            if (doc.Root.Name != PackageRelsNs + "Relationships")
            {
                _issues.Error($"{Prefix}_RELS_ROOT",
                    $"The root element must be <Relationships xmlns=\"{PackageRelsNs.NamespaceName}\"> but is <{doc.Root.Name.LocalName}> in namespace \"{doc.Root.Name.NamespaceName}\".",
                    relsPart);
                return;
            }

            var rels = new Dictionary<string, Relationship>(StringComparer.Ordinal);
            _relationships[source] = rels;

            foreach (var el in doc.Root.Elements(PackageRelsNs + "Relationship"))
            {
                var id = (string?)el.Attribute("Id");
                var type = (string?)el.Attribute("Type");
                var target = (string?)el.Attribute("Target");
                var mode = (string?)el.Attribute("TargetMode");

                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(type) || target == null)
                {
                    _issues.Error($"{Prefix}_RELATIONSHIP_INVALID", "<Relationship> requires Id, Type, and Target attributes.", relsPart, ZipPackage.LineOf(el));
                    continue;
                }

                var external = string.Equals(mode, "External", StringComparison.Ordinal);
                if (mode != null && !external && mode != "Internal")
                    _issues.Error($"{Prefix}_RELATIONSHIP_INVALID", $"TargetMode must be \"Internal\" or \"External\", not \"{mode}\".", relsPart, ZipPackage.LineOf(el));

                var resolved = external ? null : ResolveTarget(source, target);
                var rel = new Relationship(id, type, target, external, resolved, el);

                if (!rels.TryAdd(id, rel))
                {
                    _issues.Error($"{Prefix}_RELATIONSHIP_DUPLICATE_ID", $"Relationship Id \"{id}\" is used more than once.", relsPart, ZipPackage.LineOf(el));
                    continue;
                }

                if (external)
                    continue;

                if (resolved == null || !_package.Exists(resolved))
                {
                    var hint = resolved != null && _package.FindCaseInsensitive(resolved) is string nearMiss
                        ? $" A part named \"{nearMiss}\" exists; ZIP paths are case-sensitive."
                        : string.Empty;
                    _issues.Error($"{Prefix}_RELATIONSHIP_TARGET_MISSING",
                        $"Relationship \"{id}\" points to \"{target}\" (resolved to \"{resolved ?? target}\"), which does not exist in the archive.{hint}",
                        relsPart, ZipPackage.LineOf(el));
                }
            }
        }

        private Dictionary<string, Relationship>? GetRelationships(string sourcePart) =>
            _relationships.TryGetValue(sourcePart, out var rels) ? rels : null;

        /// <summary>
        /// Every r:id style attribute in a part must name a relationship declared in that part's .rels file.
        /// </summary>
        private void ValidateRelationshipIdReferences(string part)
        {
            if (part.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) || part == "[Content_Types].xml")
                return;

            var doc = _package.LoadXml(part);
            if (doc?.Root == null)
                return;

            var rels = GetRelationships(part);
            foreach (var attr in doc.Root.DescendantsAndSelf().Attributes())
            {
                if (!OfficeRelsNamespaces.Contains(attr.Name.NamespaceName) || attr.Value.Length == 0)
                    continue;
                if (rels == null || !rels.ContainsKey(attr.Value))
                    _issues.Error($"{Prefix}_RELATIONSHIP_ID_MISSING",
                        $"Attribute r:{attr.Name.LocalName}=\"{attr.Value}\" refers to a relationship that is not declared in {RelsPathFor(part)}.",
                        part, ZipPackage.LineOf(attr));
            }
        }

        // ---------------------------------------------------------------- Workbook

        private void ValidateWorkbook(string workbookPart)
        {
            var contentType = GetContentType(workbookPart);
            if (contentType != null && !WorkbookContentTypes.Contains(contentType))
            {
                var hint = contentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", StringComparison.OrdinalIgnoreCase)
                    ? " That is the MIME type of the whole .xlsx file; the workbook part needs the \".sheet.main+xml\" content type."
                    : string.Empty;
                _issues.Error($"{Prefix}_CONTENT_TYPE_WRONG",
                    $"The workbook part has content type \"{contentType}\" but must be \"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\".{hint}",
                    "[Content_Types].xml", $"part /{workbookPart}");
            }

            var doc = _package.LoadXml(workbookPart);
            if (doc?.Root == null)
                return;

            var root = doc.Root;
            if (root.Name.LocalName != "workbook" || (root.Name.Namespace != TransitionalNs && root.Name.Namespace != StrictNs))
            {
                _issues.Error($"{Prefix}_WORKBOOK_ROOT",
                    $"The root element must be <workbook xmlns=\"{TransitionalNs.NamespaceName}\"> but is <{root.Name.LocalName}> in namespace \"{root.Name.NamespaceName}\".",
                    workbookPart);
                return;
            }

            _ns = root.Name.Namespace;
            if (_ns == StrictNs)
                _issues.Warning($"{Prefix}_STRICT_CONFORMANCE",
                    "The workbook uses Strict Open XML. Excel can open it, but many other applications and libraries only support the Transitional namespace.",
                    workbookPart);

            CheckChildOrder(root, WorkbookSequence, RepeatableWorkbookElements, workbookPart);

            var sheetsElements = root.Elements(_ns + "sheets").ToList();
            if (sheetsElements.Count == 0)
            {
                _issues.Error($"{Prefix}_NO_SHEETS", "The workbook has no <sheets> element. A workbook must contain at least one sheet.", workbookPart);
                return;
            }

            var sheetElements = sheetsElements[0].Elements(_ns + "sheet").ToList();
            if (sheetElements.Count == 0)
            {
                _issues.Error($"{Prefix}_NO_SHEETS", "The <sheets> element is empty. A workbook must contain at least one <sheet>.", workbookPart, ZipPackage.LineOf(sheetsElements[0]));
                return;
            }

            foreach (var sheet in sheetElements)
                if ((string?)sheet.Attribute("name") is string sheetName)
                    _sheetNames.Add(sheetName);

            var workbookRels = GetRelationships(workbookPart);
            if (workbookRels == null)
            {
                _issues.Error($"{Prefix}_WORKBOOK_RELS_MISSING",
                    $"The workbook has no relationships part ({RelsPathFor(workbookPart)}), so its sheets can't be located.", workbookPart);
                return;
            }

            var styles = LoadStyles(workbookRels);
            var sharedStringCount = LoadSharedStrings(workbookRels);

            var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sheetIds = new HashSet<uint>();
            var visibleSheets = 0;
            var worksheets = new List<(string Name, string Part)>();

            foreach (var sheet in sheetElements)
            {
                var name = (string?)sheet.Attribute("name");
                var sheetIdText = (string?)sheet.Attribute("sheetId");
                var rId = sheet.Attributes().FirstOrDefault(a => a.Name.LocalName == "id" && OfficeRelsNamespaces.Contains(a.Name.NamespaceName))?.Value;
                var state = (string?)sheet.Attribute("state") ?? "visible";
                var location = ZipPackage.At($"sheet \"{name}\"", sheet);

                if (string.IsNullOrEmpty(name))
                    _issues.Error($"{Prefix}_SHEET_NAME_INVALID", "A <sheet> is missing its name attribute.", workbookPart, ZipPackage.LineOf(sheet));
                else
                {
                    if (ValidateSheetName(name) is string problem)
                        _issues.Error($"{Prefix}_SHEET_NAME_INVALID", $"Sheet name \"{name}\" {problem}", workbookPart, location);
                    if (!sheetNames.Add(name))
                        _issues.Error($"{Prefix}_SHEET_NAME_DUPLICATE", $"Sheet name \"{name}\" is used more than once. Sheet names must be unique (case-insensitive).", workbookPart, location);
                }

                if (!uint.TryParse(sheetIdText, NumberStyles.None, CultureInfo.InvariantCulture, out var sheetId) || sheetId == 0)
                    _issues.Error($"{Prefix}_SHEET_ID_INVALID", $"sheetId \"{sheetIdText}\" must be a positive integer.", workbookPart, location);
                else if (!sheetIds.Add(sheetId))
                    _issues.Error($"{Prefix}_SHEET_ID_DUPLICATE", $"sheetId {sheetId} is used by more than one sheet.", workbookPart, location);

                if (state is not ("visible" or "hidden" or "veryHidden"))
                    _issues.Error($"{Prefix}_SHEET_STATE_INVALID", $"state \"{state}\" must be \"visible\", \"hidden\", or \"veryHidden\".", workbookPart, location);
                else if (state == "visible")
                    visibleSheets++;

                if (string.IsNullOrEmpty(rId))
                {
                    _issues.Error($"{Prefix}_SHEET_RID_MISSING",
                        "The <sheet> has no r:id attribute linking it to its worksheet part. Declare xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" and add r:id.",
                        workbookPart, location);
                    continue;
                }

                if (!workbookRels.TryGetValue(rId, out var rel) || rel.ResolvedPath == null || !_package.Exists(rel.ResolvedPath))
                    continue; // Reported by the relationship checks.

                if (rel.Kind == "worksheet")
                    worksheets.Add((name ?? "", rel.ResolvedPath));
                else if (rel.Kind is not ("chartsheet" or "dialogsheet" or "macrosheet" or "xlMacrosheet" or "xlIntlMacrosheet"))
                    _issues.Error($"{Prefix}_SHEET_RELATIONSHIP_TYPE",
                        $"The sheet's relationship \"{rId}\" has type \"{rel.Type}\", which is not a sheet relationship. Use \"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\".",
                        workbookPart, location);
            }

            if (visibleSheets == 0 && sheetElements.Count > 0)
                _issues.Error($"{Prefix}_NO_VISIBLE_SHEET", "Every sheet is hidden. At least one sheet must be visible.", workbookPart);

            var activeTab = root.Element(_ns + "bookViews")?.Element(_ns + "workbookView")?.Attribute("activeTab");
            if (activeTab != null && (!int.TryParse(activeTab.Value, out var tab) || tab < 0 || tab >= sheetElements.Count))
                _issues.Warning($"{Prefix}_ACTIVE_TAB_INVALID",
                    $"activeTab=\"{activeTab.Value}\" does not refer to an existing sheet (valid values are 0 to {sheetElements.Count - 1}).",
                    workbookPart, ZipPackage.LineOf(activeTab));

            foreach (var (name, part) in worksheets)
                ValidateWorksheet(name, part, styles, sharedStringCount);
        }

        /// <summary>
        /// Returns a description of what is wrong with a sheet name, or null if it is valid.
        /// </summary>
        internal static string? ValidateSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "is blank.";
            if (name.Length > MaxSheetNameLength)
                return $"is {name.Length} characters long; sheet names are limited to {MaxSheetNameLength} characters.";
            var invalid = name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']);
            if (invalid >= 0)
                return $"contains '{name[invalid]}'. Sheet names can't contain [ ] : * ? / or \\.";
            if (name.StartsWith('\'') || name.EndsWith('\''))
                return "starts or ends with an apostrophe, which is not allowed.";
            if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
                return "is reserved by Excel.";
            return null;
        }

        // ---------------------------------------------------------------- Styles and shared strings

        private sealed record StyleInfo(int CellXfCount);

        private StyleInfo? LoadStyles(Dictionary<string, Relationship> workbookRels)
        {
            var styleRels = workbookRels.Values.Where(r => r.Kind == "styles" && r.ResolvedPath != null && _package.Exists(r.ResolvedPath)).ToList();
            if (styleRels.Count == 0)
                return null;

            var part = styleRels[0].ResolvedPath!;
            ExpectContentType(part, StylesContentType, "styles");

            var doc = _package.LoadXml(part);
            if (doc?.Root == null)
                return null;

            var root = doc.Root;
            if (root.Name != _ns + "styleSheet")
            {
                _issues.Error($"{Prefix}_STYLES_ROOT",
                    $"The root element must be <styleSheet xmlns=\"{_ns.NamespaceName}\"> but is <{root.Name.LocalName}> in namespace \"{root.Name.NamespaceName}\".", part);
                return null;
            }

            CheckChildOrder(root, StyleSheetSequence, [], part);

            int Count(string container, string item) => root.Element(_ns + container)?.Elements(_ns + item).Count() ?? 0;

            var fontCount = Count("fonts", "font");
            var fillCount = Count("fills", "fill");
            var borderCount = Count("borders", "border");
            var cellStyleXfCount = Count("cellStyleXfs", "xf");
            var customNumFmts = root.Element(_ns + "numFmts")?.Elements(_ns + "numFmt")
                .Select(e => (int?)e.Attribute("numFmtId") ?? -1).ToHashSet() ?? new HashSet<int>();

            var cellXfs = root.Element(_ns + "cellXfs")?.Elements(_ns + "xf").ToList() ?? new List<XElement>();
            if (cellXfs.Count == 0)
                _issues.Warning($"{Prefix}_STYLES_NO_CELLXFS", "styles.xml has no <cellXfs> entries. Excel expects at least one (the default cell format).", part);

            for (int i = 0; i < cellXfs.Count; i++)
            {
                var xf = cellXfs[i];
                var where = ZipPackage.At($"cellXfs index {i}", xf);
                CheckIndex(xf, "fontId", fontCount, "fonts", part, where);
                CheckIndex(xf, "fillId", fillCount, "fills", part, where);
                CheckIndex(xf, "borderId", borderCount, "borders", part, where);
                CheckIndex(xf, "xfId", cellStyleXfCount, "cellStyleXfs", part, where);

                if (int.TryParse((string?)xf.Attribute("numFmtId"), out var numFmtId) && numFmtId >= 164 && !customNumFmts.Contains(numFmtId))
                    _issues.Error($"{Prefix}_STYLE_NUMFMT_MISSING",
                        $"numFmtId {numFmtId} is a custom format ID (164 or higher) but no <numFmt numFmtId=\"{numFmtId}\"> is defined in <numFmts>.",
                        part, where);
            }

            return new StyleInfo(cellXfs.Count);
        }

        private void CheckIndex(XElement xf, string attribute, int count, string collection, string part, string location)
        {
            var value = (string?)xf.Attribute(attribute);
            if (value == null)
                return;
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var index) || index >= count)
                _issues.Error($"{Prefix}_STYLE_INDEX_OUT_OF_RANGE",
                    $"{attribute}=\"{value}\" is out of range; <{collection}> contains {count} item(s), so valid indexes are 0 to {count - 1}.",
                    part, location);
        }

        private int? LoadSharedStrings(Dictionary<string, Relationship> workbookRels)
        {
            var rel = workbookRels.Values.FirstOrDefault(r => r.Kind == "sharedStrings" && r.ResolvedPath != null && _package.Exists(r.ResolvedPath));
            if (rel == null)
                return null;

            var part = rel.ResolvedPath!;
            ExpectContentType(part, SharedStringsContentType, "shared strings");

            var doc = _package.LoadXml(part);
            if (doc?.Root == null)
                return 0;

            if (doc.Root.Name != _ns + "sst")
            {
                _issues.Error($"{Prefix}_SHARED_STRINGS_ROOT",
                    $"The root element must be <sst xmlns=\"{_ns.NamespaceName}\"> but is <{doc.Root.Name.LocalName}> in namespace \"{doc.Root.Name.NamespaceName}\".", part);
                return 0;
            }

            var items = doc.Root.Elements(_ns + "si").ToList();
            for (int i = 0; i < items.Count; i++)
            {
                var length = items[i].Descendants(_ns + "t").Sum(t => t.Value.Length);
                if (length > MaxCellTextLength)
                    _issues.Error($"{Prefix}_TEXT_TOO_LONG",
                        $"Shared string {i} is {length:N0} characters long; Excel cells are limited to {MaxCellTextLength:N0} characters.",
                        part, ZipPackage.LineOf(items[i]));
            }

            return items.Count;
        }

        // ---------------------------------------------------------------- Worksheets

        private void ValidateWorksheet(string sheetName, string part, StyleInfo? styles, int? sharedStringCount)
        {
            ExpectContentType(part, WorksheetContentType, "worksheet");

            var doc = _package.LoadXml(part);
            if (doc?.Root == null)
                return;

            var root = doc.Root;
            if (root.Name != _ns + "worksheet")
            {
                _issues.Error($"{Prefix}_WORKSHEET_ROOT",
                    $"The root element must be <worksheet xmlns=\"{_ns.NamespaceName}\"> but is <{root.Name.LocalName}> in namespace \"{root.Name.NamespaceName}\".", part);
                return;
            }

            CheckChildOrder(root, WorksheetSequence, RepeatableWorksheetElements, part);

            var sheetData = root.Element(_ns + "sheetData");
            if (sheetData == null)
            {
                _issues.Error($"{Prefix}_SHEETDATA_MISSING", $"Worksheet \"{sheetName}\" has no <sheetData> element. It is required even when the sheet is empty.", part, sheet: sheetName);
                return;
            }

            ValidateColumns(root, part, sheetName);
            ValidateRows(sheetData, sheetName, part, styles, sharedStringCount);

            if (!sheetData.Elements(_ns + "row").Elements(_ns + "c").Any(c => c.Element(_ns + "v") != null || c.Element(_ns + "is") != null || c.Element(_ns + "f") != null))
                _issues.Warning($"{Prefix}_SHEET_EMPTY", $"Worksheet \"{sheetName}\" contains no data.", part, sheet: sheetName);

            var dimension = root.Element(_ns + "dimension")?.Attribute("ref");
            if (dimension != null && ParseRange(dimension.Value, allowSingleCell: true) == null)
                _issues.Error($"{Prefix}_REF_INVALID", $"<dimension ref=\"{dimension.Value}\"> is not a valid cell range.", part, ZipPackage.LineOf(dimension), sheetName);

            var autoFilter = root.Element(_ns + "autoFilter")?.Attribute("ref");
            if (autoFilter != null && ParseRange(autoFilter.Value, allowSingleCell: true) == null)
                _issues.Error($"{Prefix}_REF_INVALID", $"<autoFilter ref=\"{autoFilter.Value}\"> is not a valid cell range.", part, ZipPackage.LineOf(autoFilter), sheetName);

            ValidateMergeCells(root, part, sheetName);
        }

        private void ValidateColumns(XElement root, string part, string sheetName)
        {
            var lastMax = 0;
            foreach (var col in root.Elements(_ns + "cols").Elements(_ns + "col"))
            {
                var minText = (string?)col.Attribute("min");
                var maxText = (string?)col.Attribute("max");
                if (!int.TryParse(minText, out var min) || !int.TryParse(maxText, out var max) || min < 1 || max < min || max > MaxColumns)
                {
                    _issues.Error($"{Prefix}_COL_RANGE_INVALID",
                        $"<col min=\"{minText}\" max=\"{maxText}\"> is invalid. min and max are required 1-based column numbers with 1 <= min <= max <= {MaxColumns}.",
                        part, ZipPackage.LineOf(col), sheetName);
                    continue;
                }
                if (min <= lastMax)
                    _issues.Error($"{Prefix}_COL_RANGE_OVERLAP",
                        $"<col min=\"{min}\" max=\"{max}\"> overlaps or precedes an earlier <col>. Column ranges must be sorted and must not overlap.",
                        part, ZipPackage.LineOf(col), sheetName);
                lastMax = Math.Max(lastMax, max);
            }
        }

        private void ValidateRows(XElement sheetData, string sheetName, string part, StyleInfo? styles, int? sharedStringCount)
        {
            var lastRow = 0;

            foreach (var child in sheetData.Elements())
            {
                if (child.Name != _ns + "row")
                {
                    _issues.Error($"{Prefix}_UNEXPECTED_ELEMENT",
                        $"<sheetData> may only contain <row> elements, but contains <{child.Name.LocalName}>.", part, ZipPackage.LineOf(child), sheetName);
                    continue;
                }

                var rowText = (string?)child.Attribute("r");
                int rowNumber;
                if (rowText == null)
                    rowNumber = lastRow + 1;
                else if (!int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out rowNumber) || rowNumber < 1 || rowNumber > MaxRows)
                {
                    _issues.Error($"{Prefix}_ROW_NUMBER_INVALID", $"Row number r=\"{rowText}\" must be an integer from 1 to {MaxRows:N0}.", part, ZipPackage.LineOf(child), sheetName);
                    continue;
                }

                if (rowNumber <= lastRow)
                    _issues.Error($"{Prefix}_ROW_ORDER",
                        $"Row {rowNumber} appears after row {lastRow}. Rows must be in ascending order with no duplicates.", part, ZipPackage.LineOf(child), sheetName, row: rowNumber);
                lastRow = Math.Max(lastRow, rowNumber);

                ValidateCells(child, rowNumber, sheetName, part, styles, sharedStringCount);
            }
        }

        private void ValidateCells(XElement row, int rowNumber, string sheetName, string part, StyleInfo? styles, int? sharedStringCount)
        {
            var lastColumn = 0;

            foreach (var cell in row.Elements())
            {
                if (cell.Name.Namespace != _ns)
                    continue;
                if (cell.Name.LocalName == "extLst")
                    continue;
                if (cell.Name.LocalName != "c")
                {
                    _issues.Error($"{Prefix}_UNEXPECTED_ELEMENT",
                        $"<row> may only contain <c> (cell) elements, but contains <{cell.Name.LocalName}>.", part, ZipPackage.LineOf(cell), sheetName, row: rowNumber);
                    continue;
                }

                var refText = (string?)cell.Attribute("r");
                int column;
                string cellName;
                if (refText == null)
                {
                    column = lastColumn + 1;
                    cellName = $"{ColumnName(column)}{rowNumber}";
                }
                else
                {
                    var parsed = ParseCellRef(refText);
                    if (parsed == null)
                    {
                        _issues.Error($"{Prefix}_CELL_REF_INVALID",
                            $"Cell reference r=\"{refText}\" is not a valid A1-style reference (column A to XFD, row 1 to {MaxRows:N0}, no '$').",
                            part, ZipPackage.At($"row {rowNumber}", cell), sheetName, row: rowNumber);
                        continue;
                    }

                    (var cellRow, column, var lowercase) = parsed.Value;
                    cellName = refText.ToUpperInvariant();
                    if (lowercase)
                        _issues.Warning($"{Prefix}_CELL_REF_LOWERCASE", $"Cell reference \"{refText}\" should use uppercase column letters.", part, ZipPackage.LineOf(cell), sheetName, cellName, rowNumber);
                    if (cellRow != rowNumber)
                        _issues.Error($"{Prefix}_CELL_ROW_MISMATCH",
                            $"Cell {cellName} is inside <row r=\"{rowNumber}\">. A cell's row number must match its parent row.", part, ZipPackage.LineOf(cell), sheetName, cellName, rowNumber);
                }

                var location = ZipPackage.At($"{sheetName}!{cellName}", cell);

                if (column <= lastColumn)
                    _issues.Error($"{Prefix}_CELL_ORDER",
                        $"Cell {cellName} appears after column {ColumnName(lastColumn)}. Cells in a row must be in ascending column order with no duplicates.",
                        part, location, sheetName, cellName, rowNumber);
                lastColumn = Math.Max(lastColumn, column);

                CheckChildOrder(cell, CellSequence, [], part);
                ValidateCellValue(cell, part, location, sharedStringCount, sheetName, cellName, rowNumber);
                ValidateCellStyle(cell, part, location, styles, sheetName, cellName, rowNumber);
            }
        }

        private void ValidateCellValue(XElement cell, string part, string location, int? sharedStringCount, string sheet, string cellName, int row)
        {
            var type = (string?)cell.Attribute("t") ?? "n";
            var v = cell.Element(_ns + "v");
            var value = v?.Value;
            var hasFormula = cell.Element(_ns + "f") != null;

            var formula = cell.Element(_ns + "f")?.Value;
            if (formula != null && formula.Contains("#REF!", StringComparison.Ordinal))
                _issues.Error($"{Prefix}_FORMULA_BROKEN_REFERENCE",
                    $"The formula \"{Truncate(formula)}\" contains #REF!, meaning it refers to a cell, range, or sheet that does not exist.", part, location, sheet, cellName, row);
            if (!string.IsNullOrEmpty(formula))
                foreach (var missing in Services.FormulaTranslator.ReferencedSheets(formula).Where(n => !_sheetNames.Contains(n)))
                    _issues.Error($"{Prefix}_FORMULA_UNKNOWN_SHEET",
                        $"The formula \"{Truncate(formula)}\" refers to sheet \"{missing}\", which does not exist in this workbook. Existing sheets: {string.Join(", ", _sheetNames)}.",
                        part, location, sheet, cellName, row);

            if (type != "inlineStr" && cell.Element(_ns + "is") != null)
                _issues.Error($"{Prefix}_CELL_INLINE_STRING_TYPE",
                    $"The cell has an <is> inline string but t=\"{type}\". Inline strings require t=\"inlineStr\".", part, location, sheet, cellName, row);

            switch (type)
            {
                case "n":
                    if (!string.IsNullOrEmpty(value) && !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                        _issues.Error($"{Prefix}_CELL_NOT_NUMERIC",
                            $"The cell is numeric (t=\"n\" or no t attribute) but its value \"{Truncate(value)}\" is not a number. For text, use t=\"inlineStr\" with <is><t>...</t></is>, or t=\"s\" with a shared string index.",
                            part, location, sheet, cellName, row);
                    break;

                case "s":
                    if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
                        _issues.Error($"{Prefix}_SHARED_STRING_INDEX_INVALID",
                            $"The cell has t=\"s\" so <v> must be a zero-based index into sharedStrings.xml, but it is \"{Truncate(value ?? "")}\". To store text directly, use t=\"inlineStr\" with <is><t>...</t></is>.",
                            part, location, sheet, cellName, row);
                    else if (sharedStringCount == null)
                        _issues.Error($"{Prefix}_SHARED_STRINGS_MISSING",
                            "The cell refers to a shared string (t=\"s\") but the workbook has no shared strings part linked from workbook.xml.rels.",
                            part, location, sheet, cellName, row);
                    else if (index >= sharedStringCount)
                        _issues.Error($"{Prefix}_SHARED_STRING_INDEX_OUT_OF_RANGE",
                            $"Shared string index {index} is out of range; sharedStrings.xml contains {sharedStringCount} string(s).",
                            part, location, sheet, cellName, row);
                    break;

                case "b":
                    if (value is not (null or "" or "0" or "1"))
                        _issues.Error($"{Prefix}_CELL_NOT_BOOLEAN",
                            $"The cell is boolean (t=\"b\") so its value must be 0 or 1, not \"{Truncate(value)}\".", part, location, sheet, cellName, row);
                    break;

                case "inlineStr":
                    var inline = cell.Element(_ns + "is");
                    if (inline == null && !hasFormula)
                        _issues.Error($"{Prefix}_INLINE_STRING_MISSING",
                            "The cell has t=\"inlineStr\" but no <is> element. Write the text as <is><t>text</t></is>, not <v>text</v>.",
                            part, location, sheet, cellName, row);
                    else if (inline != null && inline.Descendants(_ns + "t").Sum(t => t.Value.Length) is var length && length > MaxCellTextLength)
                        _issues.Error($"{Prefix}_TEXT_TOO_LONG",
                            $"The cell text is {length:N0} characters long; Excel cells are limited to {MaxCellTextLength:N0} characters.", part, location, sheet, cellName, row);
                    break;

                case "str":
                    if (value != null && value.Length > MaxCellTextLength)
                        _issues.Error($"{Prefix}_TEXT_TOO_LONG",
                            $"The cell text is {value.Length:N0} characters long; Excel cells are limited to {MaxCellTextLength:N0} characters.", part, location, sheet, cellName, row);
                    break;

                case "e":
                    if (!string.IsNullOrEmpty(value) && !ErrorValues.Contains(value))
                        _issues.Error($"{Prefix}_CELL_ERROR_VALUE_INVALID",
                            $"The cell is an error (t=\"e\") but \"{Truncate(value)}\" is not a recognized error value such as #N/A or #DIV/0!.", part, location, sheet, cellName, row);
                    else if (!string.IsNullOrEmpty(value))
                        _issues.Add(FormulaErrorSeverity(value), $"{Prefix}_CELL_FORMULA_ERROR",
                            $"The cell contains the error value {value}.{FormulaErrorHint(value)}", part, location, sheet, cellName, row);
                    break;

                case "d":
                    if (!string.IsNullOrEmpty(value) && !IsIsoDate(value))
                        _issues.Error($"{Prefix}_CELL_NOT_DATE",
                            $"The cell is a date (t=\"d\") but \"{Truncate(value)}\" is not an ISO 8601 date such as 2024-01-31 or 2024-01-31T13:45:00.", part, location, sheet, cellName, row);
                    break;

                default:
                    _issues.Error($"{Prefix}_CELL_TYPE_INVALID",
                        $"Cell type t=\"{type}\" is not valid. Use one of: n (number), s (shared string), inlineStr, str (formula text), b (boolean), e (error), d (date).",
                        part, location, sheet, cellName, row);
                    break;
            }
        }

        private void ValidateCellStyle(XElement cell, string part, string location, StyleInfo? styles, string sheet, string cellName, int row)
        {
            var styleText = (string?)cell.Attribute("s");
            if (styleText == null)
                return;

            if (!int.TryParse(styleText, NumberStyles.None, CultureInfo.InvariantCulture, out var styleIndex))
            {
                _issues.Error($"{Prefix}_CELL_STYLE_INVALID", $"Style index s=\"{styleText}\" must be a non-negative integer.", part, location, sheet, cellName, row);
                return;
            }

            if (styleIndex == 0)
                return;

            if (styles == null)
                _issues.Error($"{Prefix}_CELL_STYLE_OUT_OF_RANGE",
                    $"The cell uses style index {styleIndex} but the workbook has no styles part (styles.xml).", part, location, sheet, cellName, row);
            else if (styleIndex >= styles.CellXfCount)
                _issues.Error($"{Prefix}_CELL_STYLE_OUT_OF_RANGE",
                    $"Style index {styleIndex} is out of range; <cellXfs> in styles.xml contains {styles.CellXfCount} format(s).", part, location, sheet, cellName, row);
        }

        private void ValidateMergeCells(XElement root, string part, string sheetName)
        {
            var merged = new List<(int Row1, int Col1, int Row2, int Col2, string Ref)>();

            foreach (var mergeCell in root.Elements(_ns + "mergeCells").Elements(_ns + "mergeCell"))
            {
                var refText = (string?)mergeCell.Attribute("ref") ?? "";
                var range = ParseRange(refText, allowSingleCell: false);
                if (range == null)
                {
                    _issues.Error($"{Prefix}_MERGE_INVALID",
                        $"<mergeCell ref=\"{refText}\"> must be a range of at least two cells, such as A1:C1.", part, ZipPackage.LineOf(mergeCell), sheetName);
                    continue;
                }

                var (r1, c1, r2, c2) = range.Value;
                var overlap = merged.FirstOrDefault(m => r1 <= m.Row2 && m.Row1 <= r2 && c1 <= m.Col2 && m.Col1 <= c2);
                if (overlap.Ref != null)
                    _issues.Error($"{Prefix}_MERGE_OVERLAP",
                        $"Merged range {refText} overlaps merged range {overlap.Ref}. Merged ranges must not overlap.", part, ZipPackage.LineOf(mergeCell), sheetName);

                merged.Add((r1, c1, r2, c2, refText));
            }
        }

        // ---------------------------------------------------------------- Shared helpers

        /// <summary>
        /// Reports unknown, duplicated, or out-of-order child elements in the spreadsheetml namespace.
        /// Elements in other namespaces (extensions, markup compatibility) are ignored.
        /// </summary>
        private void CheckChildOrder(XElement parent, string[] sequence, ICollection<string> repeatable, string part)
        {
            var lastIndex = -1;
            string? lastName = null;

            foreach (var child in parent.Elements())
            {
                if (child.Name.Namespace != _ns)
                    continue;

                var name = child.Name.LocalName;
                var index = Array.IndexOf(sequence, name);
                if (index < 0)
                {
                    _issues.Error($"{Prefix}_UNEXPECTED_ELEMENT",
                        $"<{name}> is not a valid child of <{parent.Name.LocalName}>. Allowed children, in order: {string.Join(", ", sequence)}.",
                        part, ZipPackage.LineOf(child));
                    continue;
                }

                if (index < lastIndex)
                    _issues.Error($"{Prefix}_ELEMENT_ORDER",
                        $"<{name}> must come before <{lastName}> inside <{parent.Name.LocalName}>. Required order: {string.Join(", ", sequence)}.",
                        part, ZipPackage.LineOf(child));
                else if (index == lastIndex && !repeatable.Contains(name))
                    _issues.Error($"{Prefix}_ELEMENT_DUPLICATE",
                        $"<{parent.Name.LocalName}> may contain only one <{name}> element.", part, ZipPackage.LineOf(child));

                if (index >= lastIndex)
                {
                    lastIndex = index;
                    lastName = name;
                }
            }
        }

        /// <summary>
        /// Parses an A1-style cell reference. Returns null if it is malformed or out of range.
        /// </summary>
        internal static (int Row, int Column, bool Lowercase)? ParseCellRef(string reference)
        {
            int i = 0, column = 0;
            var lowercase = false;

            while (i < reference.Length && char.IsAsciiLetter(reference[i]))
            {
                if (i == 3)
                    return null;
                lowercase |= char.IsAsciiLetterLower(reference[i]);
                column = column * 26 + (char.ToUpperInvariant(reference[i]) - 'A' + 1);
                i++;
            }

            if (i == 0 || i == reference.Length || reference[i] == '0' || reference.Length - i > 7)
                return null;

            if (!int.TryParse(reference.AsSpan(i), NumberStyles.None, CultureInfo.InvariantCulture, out var row))
                return null;

            if (row < 1 || row > MaxRows || column > MaxColumns)
                return null;

            return (row, column, lowercase);
        }

        private static (int Row1, int Col1, int Row2, int Col2)? ParseRange(string range, bool allowSingleCell)
        {
            var parts = range.Split(':');
            if (parts.Length == 1 && allowSingleCell)
            {
                var single = ParseCellRef(parts[0]);
                return single == null ? null : (single.Value.Row, single.Value.Column, single.Value.Row, single.Value.Column);
            }
            if (parts.Length != 2)
                return null;

            var start = ParseCellRef(parts[0]);
            var end = ParseCellRef(parts[1]);
            if (start == null || end == null || end.Value.Row < start.Value.Row || end.Value.Column < start.Value.Column)
                return null;
            if (!allowSingleCell && start.Value.Row == end.Value.Row && start.Value.Column == end.Value.Column)
                return null;

            return (start.Value.Row, start.Value.Column, end.Value.Row, end.Value.Column);
        }

        internal static string ColumnName(int column)
        {
            var name = "";
            while (column > 0)
            {
                column--;
                name = (char)('A' + column % 26) + name;
                column /= 26;
            }
            return name;
        }

        private static bool IsIsoDate(string value)
        {
            try
            {
                XmlConvert.ToDateTimeOffset(value);
                return true;
            }
            catch (FormatException)
            {
                return DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _)
                    || DateTime.TryParseExact(value, ["yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ss.FFFFFFF", "yyyy-MM-ddTHH:mm"],
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out _)
                    || TimeOnly.TryParseExact(value, ["HH:mm:ss", "HH:mm"], CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
            }
        }

        /// <summary>
        /// Broken references (#REF!) and unknown names (#NAME?) indicate a malformed formula; other error values
        /// (#N/A, #DIV/0!, ...) can be legitimate results of correct formulas on the current data.
        /// </summary>
        internal static ValidationSeverity FormulaErrorSeverity(string errorValue) =>
            errorValue is "#REF!" or "#NAME?" ? ValidationSeverity.Error : ValidationSeverity.Warning;

        internal static string FormulaErrorHint(string errorValue) => errorValue switch
        {
            "#REF!" => " The formula refers to a cell, range, or sheet that does not exist.",
            "#NAME?" => " The formula uses a function or defined name that does not exist.",
            _ => string.Empty
        };

        internal static string Truncate(string value) => value.Length <= 40 ? value : value[..40] + "...";
    }
}
