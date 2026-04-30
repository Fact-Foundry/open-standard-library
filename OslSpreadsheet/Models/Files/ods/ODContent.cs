using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.ods
{
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
    [XmlRoot("document-content", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", IsNullable = false)]
    public class ODContent
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

        private Body _body;

        public ODContent()
        {
            xmlns = new XmlSerializerNamespaces();
            xmlns.Add("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
            xmlns.Add("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            xmlns.Add("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            xmlns.Add("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            xmlns.Add("draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
            xmlns.Add("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
            xmlns.Add("xlink", "http://www.w3.org/1999/xlink");
            xmlns.Add("dc", "http://purl.org/dc/elements/1.1/");
            xmlns.Add("number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0");
            xmlns.Add("svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0");
            xmlns.Add("of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2");

            fontFaceDecals = new();
            automaticStyles = new();
            _body = new();
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public decimal version { get => 1.3M; set { } }

        [XmlElement("font-face-decls", ElementName = "font-face-decls", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public FontFaceDecals fontFaceDecals { get; set; }

        [XmlElement("automatic-styles", ElementName = "automatic-styles", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public AutomaticStyles automaticStyles { get; set; }

        [XmlElement("body", ElementName = "body", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public Body body { get => _body; set { } }

        [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
        public class FontFaceDecals
        {
            public FontFaceDecals()
            {
                fontFace = new() { new FontFace() };
            }

            [XmlElement("font-face", ElementName = "font-face", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
            public List<FontFace> fontFace { get; set; }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class FontFace
            {
                [XmlAttribute("name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string Name { get; set; } = "Calibri";

                [XmlAttribute("font-family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")]
                public string FontFamily { get; set; } = "Calibri";
            }
        }

        [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
        public class AutomaticStyles
        {
            public AutomaticStyles()
            {
                automaticStyles = new List<Style>()
                {
                    new Style()
                    {
                        Name = "ce1",
                        Family = "table-cell",
                        ParentStyleName = "Default",
                        DataStyleName = "N0"
                    },
                    new Style()
                    {
                        Name = "co1",
                        Family = "table-column",
                        tableColumnProperties = new ()
                    },
                    new Style()
                    {
                        Name = "ro1",
                        Family = "table-row",
                        tableRowProperties = new ()
                    }
                };
            }

            [XmlElement("style", ElementName = "style", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
            public List<Style> automaticStyles { get; set; }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class Style
            {
                [XmlAttribute("name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? Name { get; set; }

                [XmlAttribute("family", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? Family { get; set; }

                [XmlAttribute("parent-style-name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? ParentStyleName { get; set; }

                [XmlAttribute("data-style-name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? DataStyleName { get; set; }

                [XmlAttribute("master-page-name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? MasterPageName { get; set; }

                [XmlElement("table-column-properties", ElementName = "table-column-properties")]
                public TableColumnProperties? tableColumnProperties { get; set; }

                [XmlElement("table-row-properties", ElementName = "table-row-properties", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public TableRowProperties? tableRowProperties { get; set; }

                [XmlElement("table-properties", ElementName = "table-properties", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public TableProperties? tableProperties { get; set; }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class TableColumnProperties
                {
                    [XmlAttribute("break-before", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string BreakBefore { get; set; } = "auto";

                    [XmlAttribute("column-width", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? ColumnWidth { get; set; } = "1.69333333333333cm";
                }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public class TableRowProperties
                {
                    [XmlAttribute("row-height", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? RowHeight { get; set; } = "15pt";

                    [XmlAttribute("use-optimal-row-height", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? UseOptimalRowHeight { get; set; } = "true";

                    [XmlAttribute("break-before", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string BreakBefore { get; set; } = "auto";
                }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public class TableProperties
                {
                    [XmlAttribute("display", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? Display { get; set; } = "true";

                    [XmlAttribute("writing-mode", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? WritingMode { get; set; } = "lr-tb";
                }
            }
        }

        [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
        public class Body
        {
            public Body()
            {
                spreadsheet = new();
            }

            [XmlElement("spreadsheet", ElementName = "spreadsheet", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
            public Spreadsheet spreadsheet { get; set; }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class Spreadsheet
            {
                public Spreadsheet()
                {
                    calculationSettings = new();
                    Tables = new();
                }

                [XmlElement("calculation-settings", ElementName = "calculation-settings", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public CalculationSettings calculationSettings { get; set; }

                [XmlElement("table", ElementName = "table", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public List<Table> Tables { get; set; }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class CalculationSettings
                {
                    [XmlAttribute("case-sensitive", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? CaseSensitive { get; set; } = "true";

                    [XmlAttribute("search-criteria-must-apply-to-whole-cell", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? SearchCriteriaMustApplyToWholeCell { get; set; } = "true";

                    [XmlAttribute("use-wildcards", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? UseWildcards { get; set; } = "false";

                    [XmlAttribute("use-regular-expressions", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? UseRegulareExpressions { get; set; } = "false";

                    [XmlAttribute("automatic-find-labels", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? AutomaticFindLabels { get; set; } = "false";
                }
            }
        }

        [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
        public class Table
        {
            public Table()
            {
                tableColumn = new() { NumberColumnsRepeated = "16384" };
                Rows = new List<TableRow>();
            }

            [XmlAttribute("name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
            public string Name { get; set; }

            [XmlAttribute("style-name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
            public string StyleName { get; set; } = "ta1";

            [XmlElement("table-column", ElementName = "table-column", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public TableColumn tableColumn { get; set; }

            [XmlElement("table-row", ElementName = "table-row", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public List<TableRow> Rows { get; set; }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class TableColumn
            {
                [XmlAttribute("style-name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string StyleName { get; set; } = "co1";

                [XmlAttribute("number-columns-repeated", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string NumberColumnsRepeated { get; set; } = "";

                [XmlAttribute("default-cell-style-name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string DefaultCellStyleName { get; set; } = "ce1";
            }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class TableRow
            {
                public TableRow()
                {
                    Cells = new();
                }

                [XmlIgnore]
                public int RowIndex { get; set; }

                [XmlAttribute("number-rows-repeated", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string? NumberRowsRepeated { get; set; }

                [XmlAttribute("style-name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string StyleName { get; set; } = "ro1";

                [XmlElement("table-cell", ElementName = "table-cell", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public List<TableCell> Cells { get; set; }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class TableCell
                {
                    [XmlAttribute("value-type", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")] // namespace:office
                    public string? ValueType { get; set; }

                    [XmlAttribute("value", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")] // namespace:office
                    public string? NumericValue { get; set; }

                    [XmlAttribute("boolean-value", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")] // namespace:office
                    public string? BooleanValue { get; set; }

                    [XmlAttribute("number-columns-repeated", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? NumberColumnsRepeated { get; set; }

                    [XmlAttribute("style-name", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? StyleName { get; set; }

                    [XmlElement("p", ElementName = "p", Namespace = "urn:oasis:names:tc:opendocument:xmlns:text:1.0")]
                    public string? TextValue { get; set; }
                }
            }
        }
    }
}
