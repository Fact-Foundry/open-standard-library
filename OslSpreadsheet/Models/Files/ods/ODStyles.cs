using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.ods
{
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
    [XmlRoot("document-styles", ElementName = "document-styles", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", IsNullable = false)]
    public class ODStyles
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

        public ODStyles()
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
            defaultStyles = new();
            automaticStyles = new();
            masterStyles = new();
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public decimal version { get => 1.3M; set { } }

        [XmlElement("font-face-decls", ElementName = "font-face-decls", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public FontFaceDecals fontFaceDecals { get; set; }

        [XmlElement("styles", ElementName = "styles", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public DefaultStyles defaultStyles { get; set; }

        [XmlElement("automatic-styles", ElementName = "automatic-styles", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public AutomaticStyles automaticStyles { get; set; }

        [XmlElement("master-styles", ElementName = "master-styles", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public MasterStyles masterStyles { get; set; }

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
                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string Name { get; set; } = "Calibri";

                [XmlAttribute("font-family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")]
                public string FontFamily { get; set; } = "Calibri";
            }
        }

        [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
        public class DefaultStyles
        {
            public DefaultStyles()
            {
                numberStyle = new();
                style = new();
                defaultStyle = new();
            }

            [XmlElement("number-style", ElementName = "number-style", Namespace = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")]
            public NumberStyle numberStyle { get; set; }

            [XmlElement("style", ElementName = "style", Namespace = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")]
            public Style style { get; set; }

            [XmlElement("default-style", ElementName = "default-style", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
            public DefaultStyle defaultStyle { get; set; }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class NumberStyle
            {
                public NumberStyle()
                {
                    Numbers = new() { new Number() };
                }

                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Name { get; set; } = "N0";

                [XmlElement("number", ElementName = "number", Namespace = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")]
                public List<Number> Numbers { get; set; }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class Number
                {
                    [XmlAttribute("min-integer-digits", Namespace = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")] // namespace:number
                    public string MinimumIntegerDigits { get; set; } = "1";
                }
            }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class Style
            {
                public Style()
                {
                    tableCellProperties = new();
                    textProperties = new();
                }

                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Name { get; set; } = "Default";

                [XmlAttribute("family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Family { get; set; } = "table-cell";

                [XmlAttribute("data-style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string DataStyleName { get; set; } = "N0";

                [XmlElement("table-cell-properties", ElementName = "table-cell-properties", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public TableCellProperties tableCellProperties { get; set; }

                [XmlElement("text-properties", ElementName = "text-properties", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public TextProperties textProperties { get; set; }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class TableCellProperties
                {
                    [XmlAttribute("vertical-align", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string VerticalAlign { get; set; } = "automatic";

                    [XmlAttribute("background-color", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string BackgroundColor { get; set; } = "transparent";
                }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class TextProperties
                {
                    [XmlAttribute("color", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string Color { get; set; } = "#000000";

                    [XmlAttribute("font-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontName { get; set; } = "Calibri";

                    [XmlAttribute("font-name-asian", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontNameAsian { get; set; } = "Calibri";

                    [XmlAttribute("font-name-complex", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontNameComplex { get; set; } = "Calibri";

                    [XmlAttribute("font-size", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string FontSize { get; set; } = "11pt";

                    [XmlAttribute("font-size-asian", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontSizeAsian { get; set; } = "11pt";

                    [XmlAttribute("font-size-complex", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontSizeCopmlex { get; set; } = "11pt";
                }
            }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class DefaultStyle
            {
                public DefaultStyle()
                {
                    graphicProperties = new();
                }

                [XmlAttribute("family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Family { get; set; } = "graphic";

                [XmlElement("graphic-properties", ElementName = "graphic-properties", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public GraphicProperties graphicProperties { get; set; }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class GraphicProperties
                {
                    [XmlAttribute("fill", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string Fill { get; set; } = "solid";

                    [XmlAttribute("fill-color", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string FillColor { get; set; } = "#4472c4";

                    [XmlAttribute("opacity", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string Opacity { get; set; } = "100%";

                    [XmlAttribute("stroke", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string Stroke { get; set; } = "solid";

                    [XmlAttribute("stroke-width", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")] // namespace:svg
                    public string StrokeWidth { get; set; } = "0.01389in";

                    [XmlAttribute("stroke-color", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")] // namespace:svg
                    public string StrokeColor { get; set; } = "#2f528f";

                    [XmlAttribute("stroke-opacity", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")] // namespace:svg
                    public string StrokeOpacity { get; set; } = "100%";

                    [XmlAttribute("stroke-linejoin", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string StrokeLineJoin { get; set; } = "miter";

                    [XmlAttribute("stroke-linecap", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")] // namespace:svg
                    public string StrokeLinecap { get; set; } = "butt";
                }
            }
        }

        [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
        public class AutomaticStyles
        {
            public AutomaticStyles()
            {
                pageLayout = new();
            }

            [XmlElement("page-layout", ElementName = "page-layout", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
            public List<PageLayout> pageLayout { get; set; }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class PageLayout
            {
                public PageLayout()
                {
                    pageLayoutProperties = new();
                    headerStyle = new();
                    footerStyle = new();
                }

                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Name { get; set; } = "pm1";

                [XmlElement("page-layout-properties", ElementName = "page-layout-properties", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public PageLayoutProperties pageLayoutProperties { get; set; }

                [XmlElement("header-style", ElementName = "header-style", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public HeaderStyle headerStyle { get; set; }

                [XmlElement("footer-style", ElementName = "footer-style", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public FooterStyle footerStyle { get; set; }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class PageLayoutProperties
                {
                    [XmlAttribute("margin-top", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string MarginTop { get; set; } = "0.3in";

                    [XmlAttribute("margin-bottom", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string MarginBottom { get; set; } = "0.3in";

                    [XmlAttribute("margin-left", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string MarginLeft { get; set; } = "0.7in";

                    [XmlAttribute("margin-right", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string MarginRight { get; set; } = "0.7in";

                    [XmlAttribute("table-centering", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string TableCentering { get; set; } = "none";

                    [XmlAttribute("print", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string Print { get; set; } = "objects charts drawings";
                }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class HeaderStyle
                {
                    public HeaderStyle()
                    {
                        headerFooterProperties = new HeaderFooterProperties()
                        {
                            MinHeight = "0.45in",
                            MarginLeft = "0.7in",
                            MarginRight = "0.7in",
                            MarginBottom = "0in"
                        };
                    }

                    [XmlElement("header-footer-properties", ElementName = "header-footer-properties", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                    public HeaderFooterProperties headerFooterProperties { get; set; }
                }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class FooterStyle
                {
                    public FooterStyle()
                    {
                        headerFooterProperties = new HeaderFooterProperties()
                        {
                            MinHeight = "0.45in",
                            MarginLeft = "0.7in",
                            MarginRight = "0.7in",
                            MarginTop = "0in"
                        };
                    }

                    [XmlElement("header-footer-properties", ElementName = "header-footer-properties", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                    public HeaderFooterProperties headerFooterProperties { get; set; }
                }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class HeaderFooterProperties
                {
                    [XmlAttribute("min-height", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string? MinHeight { get; set; }

                    [XmlAttribute("margin-left", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string? MarginLeft { get; set; }

                    [XmlAttribute("margin-right", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string? MarginRight { get; set; }

                    [XmlAttribute("margin-bottom", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string? MarginBottom { get; set; }

                    [XmlAttribute("margin-top", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string? MarginTop { get; set; }
                }
            }
        }

        [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
        public class MasterStyles
        {
            public MasterStyles()
            {
                masterPage = new();
            }

            [XmlElement("master-page", ElementName = "master-page", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
            public List<MasterPage> masterPage { get; set; }

            [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public class MasterPage
            {
                public MasterPage()
                {
                    HeaderPageStyle = new();
                    HeaderLeftPageStyle = new()
                    {
                        Display = "false"
                    };
                    HeaderFirstPageStyle = new();
                    FooterPageStyle = new();
                    FooterLeftPageStyle = new()
                    {
                        Display = "false"
                    };
                    FooterFirstPageStyle = new();
                }

                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Name { get; set; } = "mp1";

                [XmlAttribute("page-layout-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string PageLayoutName { get; set; } = "pm1";

                [XmlElement("header", ElementName = "header", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public PageStyle HeaderPageStyle { get; set; }

                [XmlElement("header-left", ElementName = "header-left", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public PageStyle HeaderLeftPageStyle { get; set; }

                [XmlElement("header-first", ElementName = "header-first", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public PageStyle HeaderFirstPageStyle { get; set; }

                [XmlElement("footer", ElementName = "footer", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public PageStyle FooterPageStyle { get; set; }

                [XmlElement("footer-left", ElementName = "footer-left", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public PageStyle FooterLeftPageStyle { get; set; }

                [XmlElement("footer-first", ElementName = "footer-first", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public PageStyle FooterFirstPageStyle { get; set; }

                [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public class PageStyle
                {
                    [XmlAttribute("display", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? Display { get; set; }
                }
            }
        }
    }
}
