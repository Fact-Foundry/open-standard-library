namespace OslSpreadsheet.Services
{
    /// <summary>
    /// Normalizes the color strings accepted by <c>CellStyle</c> and <c>CellBorder</c> before they are written into XML.
    /// </summary>
    internal static class Colors
    {
        /// <summary>
        /// Returns the color as "#RRGGBB" (uppercase) when it is six hex digits with an optional leading '#'; otherwise null.
        /// Only this form is accepted, so a color can never carry other characters into a generated file.
        /// </summary>
        internal static string? Normalize(string? color)
        {
            if (string.IsNullOrEmpty(color))
                return null;

            var hex = color.StartsWith('#') ? color[1..] : color;
            if (hex.Length != 6 || !hex.All(char.IsAsciiHexDigit))
                return null;

            return "#" + hex.ToUpperInvariant();
        }

        /// <summary>
        /// The ARGB form Excel uses in styles.xml ("FFRRGGBB"). The color must already be valid.
        /// </summary>
        internal static string ToArgb(string color) => "FF" + Normalize(color)![1..];
    }
}
