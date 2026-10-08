namespace OslSpreadsheet.Models
{
    public class oWorkbook
    {
        public oWorkbook()
        {
            Sheets = new List<oSpreadsheet>();
        }

        private int _nextSheetIndex
        {
            get => Sheets.Any() ? Sheets.Max(x => x.Index) + 1 : 1;
        }

        /// <summary>
        /// Application that generated this file
        /// </summary>
        public string Generator { get; set; } = AssemblyInfo.AppName;

        public string InitialCreator { get; set; } = "";

        public string Creator { get; set; } = "";

        public string CreationDate { get; set; } = "";

        public ColumnDelimeter ColumnDelimeter { get; set; } = ColumnDelimeter.Comma;

        /// <summary>
        /// Character encoding used when reading or writing delimited files (CSV, TSV, etc.).
        /// Does not apply to ODS or XLSX formats, which require UTF-8 per their specifications.
        /// </summary>
        public FileEncoding FileEncoding { get; set; } = FileEncoding.UTF8;

        public List<oSpreadsheet> Sheets { get; set; }

        /// <summary>
        /// Checks whether this workbook can be written as a valid file of the given format, without generating it.
        /// Reports empty workbooks, invalid or duplicate sheet names, values that don't match their <see cref="CellValueType"/>,
        /// out-of-range cell positions, control characters, over-long text, and formulas that refer to missing sheets.
        /// The Generate methods run the same checks and throw <see cref="Validation.InvalidWorkbookException"/> on errors.
        /// </summary>
        public Validation.ValidationResult Validate(Validation.ValidationFileFormat format = Validation.ValidationFileFormat.Xlsx) =>
            Validation.WorkbookValidator.Validate(this, format);

        public oSpreadsheet AddSheet()
        {
            var index = _nextSheetIndex;

            return _AddSheet(new oSpreadsheet(index, string.Format("Sheet{0}", index)));
        }

        public oSpreadsheet AddSheet(string name)
        {
            var index = _nextSheetIndex;

            return _AddSheet(new oSpreadsheet(index, name));
        }

        public async Task<oSpreadsheet> AddSheetAsync()
        {
            var index = _nextSheetIndex;

            return await Task.Run(() => _AddSheet(new oSpreadsheet(index, string.Format("Sheet{0}", index))));
        }

        public async Task<oSpreadsheet> AddSheetAsync(string name)
        {
            var index = _nextSheetIndex;

            return await Task.Run(() => _AddSheet(new oSpreadsheet(index, name)));
        }

        private oSpreadsheet _AddSheet(oSpreadsheet sheet)
        {
            try
            {
                lock (this)
                {
                    var index = Sheets.FindIndex(x => x.Index == sheet.Index);

                    if (index == -1)
                        Sheets.Add(sheet);
                    else
                        Sheets[index] = sheet;

                    return sheet;
                }
            }
            catch
            {
                throw new Exception("There was an error adding a new worksheet to the workbook");
            }
        }
    }
}
