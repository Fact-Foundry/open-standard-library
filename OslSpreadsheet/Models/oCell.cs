namespace OslSpreadsheet.Models
{
    public class oCell
    {
        private readonly int _rowIndex;

        private readonly int _columnIndex;

        public oCell(int row, int column)
        {
            _columnIndex = column;
            _rowIndex = row;
        }

        public int Row { get => _rowIndex; }

        public int Column { get => _columnIndex; }

        public CellValueType ValueType { get; set; } = CellValueType.String;

        public string Value { get; set; } = "";

        public string? Formula { get; set; }

        public CellStyle? Style { get; set; }

        /// <summary>
        /// Returns the value as a spreadsheet application would display it, by applying <see cref="CellStyle.NumberFormat"/>:
        /// <c>1234.5</c> with <c>"$"#,##0.00</c> becomes <c>$1,234.50</c>, and <c>2021-03-15</c> with <c>mm/dd/yyyy</c>
        /// becomes <c>03/15/2021</c>. Without a format, DateTime cells use <c>yyyy-MM-dd</c> (plus <c>HH:mm:ss</c> when there
        /// is a time), Boolean cells show <c>TRUE</c>/<c>FALSE</c>, and other cells return <see cref="Value"/>.
        /// Codes with several sections use the positive, negative, zero, or text section as Excel does, and padding at
        /// the ends is trimmed. Formats the library can't apply (conditions, scientific notation, fractions, elapsed time)
        /// return <see cref="Value"/> unchanged, as do values that don't match their type.
        /// </summary>
        /// <param name="culture">Supplies the decimal and group separators, month and day names, and AM/PM designators.
        /// Defaults to the invariant culture, so the result doesn't depend on the machine's locale.</param>
        public string GetFormattedValue(System.Globalization.CultureInfo? culture = null) =>
            Services.CellValueFormatter.Format(this, culture);
    }

    public enum CellValueType
    {
        String,
        Float,
        Boolean,
        DateTime,
        Int64
    }
}
