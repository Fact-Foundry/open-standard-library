namespace OslSpreadsheet.Models
{
    public class CellStyle
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string? FontColor { get; set; }
        public string? BackgroundColor { get; set; }
        public string? FontName { get; set; }
        public double? FontSize { get; set; }
        public bool WrapText { get; set; }

        /// <summary>
        /// Number format as an Excel format code, e.g. "#,##0.00", "0%", "\"$\"#,##0.00", "yyyy-mm-dd", "mmm d, yyyy", "h:mm AM/PM".
        /// Written to XLSX as-is and translated to an ODS data style. Null uses the application default
        /// (DateTime cells default to yyyy-mm-dd or yyyy-mm-dd hh:mm:ss).
        /// </summary>
        public string? NumberFormat { get; set; }

        public CellBorder? BorderTop { get; set; }
        public CellBorder? BorderBottom { get; set; }
        public CellBorder? BorderLeft { get; set; }
        public CellBorder? BorderRight { get; set; }

        /// <summary>
        /// True when no property is set, i.e. the style would have no visible effect.
        /// </summary>
        public bool IsDefault =>
            !Bold && !Italic && !Underline && FontColor == null && BackgroundColor == null && FontName == null && FontSize == null
            && !WrapText && NumberFormat == null && BorderTop == null && BorderBottom == null && BorderLeft == null && BorderRight == null;

        /// <summary>
        /// Returns a copy of this style. Borders are copied too, so the copy can be changed independently.
        /// </summary>
        public CellStyle Clone() => new()
        {
            Bold = Bold, Italic = Italic, Underline = Underline, FontColor = FontColor, BackgroundColor = BackgroundColor,
            FontName = FontName, FontSize = FontSize, WrapText = WrapText, NumberFormat = NumberFormat,
            BorderTop = BorderTop?.Clone(), BorderBottom = BorderBottom?.Clone(), BorderLeft = BorderLeft?.Clone(), BorderRight = BorderRight?.Clone()
        };
    }

    public class CellBorder
    {
        public BorderStyle Style { get; set; } = BorderStyle.Thin;
        public string? Color { get; set; }

        public CellBorder Clone() => new() { Style = Style, Color = Color };
    }

    public enum BorderStyle
    {
        None,
        Thin,
        Medium,
        Thick
    }
}
