namespace OslSpreadsheet.Models
{
    public class oSpreadsheet
    {
        private readonly int _index;

        private readonly List<oCell> _cells;

        private readonly Dictionary<int, double> _columnWidths;

        // Position index for O(1) cell lookup. Rebuilt if Cells was modified directly (detected by a count mismatch).
        private readonly Dictionary<(int Row, int Column), oCell> _cellIndex;

        public oSpreadsheet(int index, string name)
        {
            _index = index;

            _cells = new();

            _columnWidths = new();

            _cellIndex = new();

            SheetName = name;
        }

        public string SheetName { get; set; }

        public List<oCell> Cells { get => _cells; }

        public int Index { get => _index; }

        public int FreezeRows { get; set; }

        public int FreezeColumns { get; set; }

        /// <summary>
        /// Indicates whether the first row contains column headers.
        /// When true, HeaderNames and GetColumn(string) become available.
        /// </summary>
        public bool HasHeaderRow { get; set; }

        /// <summary>
        /// Returns the values from the first row when HasHeaderRow is true; otherwise an empty list.
        /// </summary>
        public List<string> HeaderNames
        {
            get
            {
                if (!HasHeaderRow || !_cells.Any()) return new List<string>();
                return GetRow(1).Select(c => c.Value).ToList();
            }
        }

        public (int StartRow, int StartCol, int EndRow, int EndCol)? AutoFilterRange { get; set; }

        public Dictionary<int, double> ColumnWidths { get => _columnWidths; }

        public void SetColumnWidth(int column, double width)
        {
            _columnWidths[column] = width;
        }

        public void SetAutoFilter()
        {
            if (!_cells.Any()) return;
            AutoFilterRange = (1, 1, RowCount, ColumnCount);
        }

        public void SetAutoFilter(int startRow, int startCol, int endRow, int endCol)
        {
            AutoFilterRange = (startRow, startCol, endRow, endCol);
        }

        public void AutoFitColumns(double minWidth = 8, double maxWidth = 100)
        {
            if (!_cells.Any()) return;

            var maxLengths = new int[ColumnCount + 1];
            foreach (var cell in _cells)
                maxLengths[cell.Column] = Math.Max(maxLengths[cell.Column], cell.Value.Length);

            for (int c = 1; c < maxLengths.Length; c++)
                _columnWidths[c] = Math.Clamp(maxLengths[c] + 2, minWidth, maxWidth);
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
                    EnsureIndex();

                    if (_cellIndex.TryGetValue((cell.Row, cell.Column), out var existing))
                        _cells[_cells.IndexOf(existing)] = cell;
                    else
                        _cells.Add(cell);

                    _cellIndex[(cell.Row, cell.Column)] = cell;

                    return cell;
                }
            }
            catch
            {
                throw new Exception("There was an error adding a new cell to the spreadsheet");
            }
        }

        /// <summary>
        /// Rebuilds the position index when Cells has been changed directly rather than through AddCell.
        /// </summary>
        private void EnsureIndex()
        {
            if (_cellIndex.Count == _cells.Count) return;

            _cellIndex.Clear();
            foreach (var cell in _cells)
                _cellIndex[(cell.Row, cell.Column)] = cell;
        }

        /// <summary>
        /// Returns the cell at the given position, or null if there is none.
        /// </summary>
        public oCell? GetCell(int row, int column)
        {
            lock (this)
            {
                EnsureIndex();
                return _cellIndex.TryGetValue((row, column), out var cell) ? cell : null;
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
        /// Returns all data cells in the column matching the given header name (excludes the header row itself).
        /// Requires HasHeaderRow to be true.
        /// </summary>
        /// <param name="headerName"></param>
        /// <returns></returns>
        public List<oCell> GetColumn(string headerName)
        {
            if (!HasHeaderRow) return new List<oCell>();
            var headerCell = GetRow(1).FirstOrDefault(c => c.Value == headerName);
            if (headerCell == null) return new List<oCell>();
            return _cells.Where(c => c.Column == headerCell.Column && c.Row > 1).OrderBy(c => c.Row).ToList();
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
