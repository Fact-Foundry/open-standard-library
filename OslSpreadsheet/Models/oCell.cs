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
    }

    public enum CellValueType
    {
        String,
        Float,
        Boolean
    }
}
