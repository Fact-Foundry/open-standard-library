namespace OslSpreadsheet.Models
{
    public class oSpreadsheet
    {
        private readonly int _index;

        private readonly List<oCell> _cells;

        private readonly Dictionary<int, double> _columnWidths;

        public oSpreadsheet(int index, string name)
        {
            _index = index;

            _cells = new();

            _columnWidths = new();

            SheetName = name;
        }

        public string SheetName { get; set; }

        public List<oCell> Cells { get => _cells; }

        public int Index { get => _index; }

        public int FreezeRows { get; set; }

        public int FreezeColumns { get; set; }

        public Dictionary<int, double> ColumnWidths { get => _columnWidths; }

        public void SetColumnWidth(int column, double width)
        {
            _columnWidths[column] = width;
        }

        public void AutoFitColumns(double minWidth = 8, double maxWidth = 100)
        {
            if (!_cells.Any()) return;

            for (int c = 1; c <= ColumnCount; c++)
            {
                var maxLength = _cells
                    .Where(x => x.Column == c)
                    .Select(x => x.Value.Length)
                    .DefaultIfEmpty(0)
                    .Max();

                var width = Math.Clamp(maxLength + 2, minWidth, maxWidth);
                _columnWidths[c] = width;
            }
        }

        public oCell AddCell(int row, int column)
        {
            return _AddCell(new oCell(row, column));
        }

        public oCell AddCell(int row, int column, string value)
        {
            return _AddCell(new oCell(row, column)
            {
                Value = value
            });
        }

        public oCell AddCell(int row, int column, string value, CellValueType valueType)
        {
            return _AddCell(new oCell(row, column)
            {
                Value = value,
                ValueType = valueType
            });
        }

        public async Task<oCell> AddCellAsync(int row, int column)
        {
            return await Task.Run(() => AddCell(row, column));
        }

        public async Task<oCell> AddCellAsync(int row, int column, string value)
        {
            return await Task.Run(() => AddCell(row, column, value));
        }

        public async Task<oCell> AddCellAsync(int row, int column, string value, CellValueType valueType)
        {
            return await Task.Run(() => AddCell(row, column, value, valueType));
        }

        private oCell _AddCell(oCell cell)
        {
            try
            {
                lock (this)
                {
                    var index = Cells.FindIndex(x => x.Row == cell.Row && x.Column == cell.Column);

                    if (index == -1)
                        Cells.Add(cell);
                    else
                        Cells[index] = cell;

                    return cell;
                }
            }
            catch
            {
                throw new Exception("There was an error adding a new cell to the spreadsheet");
            }
        }

        /// <summary>
        /// Get a count of columns in a spreadsheet
        /// </summary>
        public int ColumnCount { get => _cells.Any() ? _cells.Max(x => x.Column) : 0; }

        /// <summary>
        /// Get a count of rows in the spreadsheet
        /// </summary>
        public int RowCount { get => _cells.Any() ? _cells.Max(x => x.Row) : 0; }

        /// <summary>
        /// Get all cells from a particular row
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public List<oCell> GetRow(int index)
        {
            return _cells.Where(x => x.Row == index).OrderBy(x => x.Column).ToList();
        }

        /// <summary>
        /// Converts a spreadsheet to a 2D array
        /// </summary>
        /// <returns></returns>
        public string[,] ToArray()
        {
            var retVal = new string[RowCount, ColumnCount];

            foreach(var c in _cells)
            {
                retVal[c.Row - 1, c.Column - 1] = c.Value;
            }

            return retVal;
        }
    }
}
