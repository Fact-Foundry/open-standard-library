namespace OslSpreadsheet.Models
{
    public static class oSpreadsheetExtensions
    {
        private static readonly DateTime UnixEpoch = new(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        /// <summary>
        /// Sets the cell's value to a number and marks it as <see cref="CellValueType.Float"/>.
        /// The value is written with a period as the decimal separator regardless of the current culture.
        /// </summary>
        public static oCell AsFloat(this oCell cell, double value)
        {
            cell.Value = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            cell.ValueType = CellValueType.Float;

            return cell;
        }

        /// <summary>
        /// Converts the cell's value from Unix epoch seconds to an ISO 8601 DateTime.
        /// </summary>
        public static oCell FromEpochSeconds(this oCell cell)
        {
            if (long.TryParse(cell.Value, out long seconds))
            {
                cell.Value = UnixEpoch.AddSeconds(seconds).ToString("yyyy-MM-ddTHH:mm:ss");
                cell.ValueType = CellValueType.DateTime;
            }
            return cell;
        }

        /// <summary>
        /// Converts the cell's value from Unix epoch milliseconds to an ISO 8601 DateTime.
        /// </summary>
        public static oCell FromEpochMilliseconds(this oCell cell)
        {
            if (long.TryParse(cell.Value, out long milliseconds))
            {
                cell.Value = UnixEpoch.AddMilliseconds(milliseconds).ToString("yyyy-MM-ddTHH:mm:ss");
                cell.ValueType = CellValueType.DateTime;
            }
            return cell;
        }

        /// <summary>
        /// Converts the cell's DateTime value to Unix epoch seconds.
        /// Values are treated as UTC.
        /// </summary>
        public static oCell ToEpochSeconds(this oCell cell)
        {
            if (DateTime.TryParse(cell.Value, System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal, out var dt))
            {
                cell.Value = ((long)(dt - UnixEpoch).TotalSeconds).ToString();
                cell.ValueType = CellValueType.Int64;
            }
            return cell;
        }

        /// <summary>
        /// Converts the cell's DateTime value to Unix epoch milliseconds.
        /// Values are treated as UTC.
        /// </summary>
        public static oCell ToEpochMilliseconds(this oCell cell)
        {
            if (DateTime.TryParse(cell.Value, System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal, out var dt))
            {
                cell.Value = ((long)(dt - UnixEpoch).TotalMilliseconds).ToString();
                cell.ValueType = CellValueType.Int64;
            }
            return cell;
        }
    }
}
