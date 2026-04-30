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
        public CellBorder? BorderTop { get; set; }
        public CellBorder? BorderBottom { get; set; }
        public CellBorder? BorderLeft { get; set; }
        public CellBorder? BorderRight { get; set; }
    }

    public class CellBorder
    {
        public BorderStyle Style { get; set; } = BorderStyle.Thin;
        public string? Color { get; set; }
    }

    public enum BorderStyle
    {
        None,
        Thin,
        Medium,
        Thick
    }
}
