using System.Globalization;
using System.Text;
using OslSpreadsheet.Models;
using static OslSpreadsheet.Services.NumberFormatTranslator;

namespace OslSpreadsheet.Services
{
    /// <summary>
    /// Renders a cell's value the way a spreadsheet application displays it, by applying the Excel format code in
    /// <see cref="CellStyle.NumberFormat"/>. Supports the same subset as the ODS translation; anything else returns the raw value.
    /// </summary>
    internal static class CellValueFormatter
    {
        internal static string Format(oCell cell, CultureInfo? culture)
        {
            culture ??= CultureInfo.InvariantCulture;
            var value = cell.Value ?? "";
            if (value.Length == 0) return "";

            // Spreadsheet applications show booleans as TRUE/FALSE whatever the format
            if (cell.ValueType == CellValueType.Boolean)
                return value.Trim().ToLowerInvariant() switch
                {
                    "true" or "1" => "TRUE",
                    "false" or "0" => "FALSE",
                    _ => value
                };

            var code = cell.Style?.NumberFormat;
            if (string.IsNullOrWhiteSpace(code) || code.Trim().Equals("General", StringComparison.OrdinalIgnoreCase))
                return FormatDefault(cell.ValueType, value);

            if (HasCondition(code))
                return value;

            // A section containing @ is the text section; Excel also treats a fourth section as one
            var sections = SplitSections(code);
            string? textSection = null;
            if (ContainsTextPlaceholder(sections[^1]))
            {
                textSection = sections[^1];
                sections.RemoveAt(sections.Count - 1);
            }
            else if (sections.Count >= 4)
            {
                textSection = sections[3];
                sections.RemoveRange(3, sections.Count - 3);
            }

            if (cell.ValueType == CellValueType.String)
                return textSection == null ? value : TrimPadding(RenderText(textSection, value));

            if (sections.Count == 0)
                return FormatDefault(cell.ValueType, value);

            var first = Parse(MarkPadding(sections[0]));
            var result = first.Kind switch
            {
                Kind.Date or Kind.Time => ToDate(cell.ValueType, value) is DateTime date ? FormatDate(first, date, culture) : null,
                _ => ToNumber(cell.ValueType, value) is decimal number ? FormatNumber(sections, number, culture) : null
            };
            return result == null ? value : TrimPadding(result);
        }

        /// <summary>
        /// Picks the section for the number's sign, as Excel does: positive and zero; negative; zero. The negative
        /// section gets the absolute value, so it shows a sign only if it contains one, as in <c>(#,##0.00)</c>.
        /// </summary>
        private static string? FormatNumber(List<string> sections, decimal number, CultureInfo culture)
        {
            string section;
            bool signed = false;
            if (number < 0 && sections.Count >= 2) { section = sections[1]; number = -number; }
            else if (number == 0 && sections.Count >= 3) section = sections[2];
            else { section = sections[0]; signed = true; }

            var spec = Parse(MarkPadding(section));
            return spec.Kind switch
            {
                Kind.Number or Kind.Percentage or Kind.Currency => FormatNumber(spec, number, signed, culture),
                // A section with no digits, such as "-" or "n/a", shows only its literal text
                Kind.Unsupported when !HasUnsupportedToken(section) => string.Concat(spec.DateParts.Select(p => p.Text)),
                Kind.Date or Kind.Time when number >= -657434 && number < 2958466 => FormatDate(spec, DateTime.FromOADate((double)number), culture),
                _ => null
            };
        }

        private static string FormatDefault(CellValueType type, string value)
        {
            if (type != CellValueType.DateTime || !TryParseDate(value, out var date))
                return value;
            return date.TimeOfDay == TimeSpan.Zero
                ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        // Stands in for the spaces that _x padding produces, so they can be trimmed from the ends without touching literal spaces
        private const char PaddingMark = '';

        /// <summary>
        /// Calls <paramref name="visit"/> for each character outside quotes, escapes, and brackets, with its index.
        /// </summary>
        private static bool AnyUnquoted(string code, Func<string, int, bool> visit)
        {
            bool inQuotes = false;
            for (int i = 0; i < code.Length; i++)
            {
                var c = code[i];
                if (c == '"') inQuotes = !inQuotes;
                else if (inQuotes) continue;
                else if (c == '\\') i++;
                else if (visit(code, i)) return true;
                else if (c == '[') { var end = code.IndexOf(']', i + 1); i = end < 0 ? code.Length : end; }
            }
            return false;
        }

        /// <summary>
        /// Conditions such as [&gt;100] choose sections by value in ways not modelled here, so codes with them aren't formatted.
        /// </summary>
        private static bool HasCondition(string code) =>
            AnyUnquoted(code, (s, i) => s[i] == '[' && i + 1 < s.Length && s[i + 1] is '<' or '>' or '=');

        private static bool ContainsTextPlaceholder(string section) => AnyUnquoted(section, (s, i) => s[i] == '@');

        /// <summary>
        /// True when a section the parser rejected has digit placeholders or elapsed time, rather than being literal text only.
        /// </summary>
        private static bool HasUnsupportedToken(string section) =>
            AnyUnquoted(section, (s, i) => s[i] is '#' or '0' or '?'
                || (s[i] == '[' && i + 1 < s.Length && "hHmMsS".Contains(s[i + 1])));

        /// <summary>
        /// Replaces each _x padding (the width of character x) with <see cref="PaddingMark"/>.
        /// </summary>
        private static string MarkPadding(string section)
        {
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < section.Length; i++)
            {
                var c = section[i];
                if (c == '"') inQuotes = !inQuotes;
                if (!inQuotes && c == '_' && i + 1 < section.Length) { sb.Append(PaddingMark); i++; continue; }
                sb.Append(c);
                if (!inQuotes && (c == '\\' || c == '*') && i + 1 < section.Length) sb.Append(section[++i]);
            }
            return sb.ToString();
        }

        private static string TrimPadding(string text) => text.Trim(PaddingMark).Replace(PaddingMark, ' ');

        /// <summary>
        /// Renders a text section: @ is the cell's text and the rest are literals, e.g. "ID-"@ gives ID-00123.
        /// </summary>
        private static string RenderText(string section, string value)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < section.Length; i++)
            {
                var c = section[i];
                switch (c)
                {
                    case '"':
                        var end = section.IndexOf('"', i + 1);
                        if (end < 0) end = section.Length;
                        sb.Append(section, i + 1, end - i - 1);
                        i = end;
                        break;
                    case '\\' when i + 1 < section.Length:
                        sb.Append(section[++i]);
                        break;
                    case '_' when i + 1 < section.Length:
                        sb.Append(PaddingMark);
                        i++;
                        break;
                    case '*' when i + 1 < section.Length:
                        i++;
                        break;
                    case '[':
                        var close = section.IndexOf(']', i + 1);
                        if (close < 0) close = section.Length;
                        var inner = section[(i + 1)..close];
                        if (inner.StartsWith('$')) sb.Append(inner[1..].Split('-')[0]);
                        i = close;
                        break;
                    case '@':
                        sb.Append(value);
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        private static bool TryParseDate(string value, out DateTime date) =>
            DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);

        private static DateTime? ToDate(CellValueType type, string value)
        {
            if (type == CellValueType.DateTime)
                return TryParseDate(value, out var date) ? date : null;

            // A number with a date format is a serial date, as in Excel
            if (type is CellValueType.Float or CellValueType.Int64
                && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial)
                && serial >= -657434 && serial < 2958466)
                return DateTime.FromOADate(serial);

            return null;
        }

        private static decimal? ToNumber(CellValueType type, string value)
        {
            if (type is CellValueType.Float or CellValueType.Int64
                && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                && Math.Abs(number) < 7.9e27)
                // decimal keeps 15 significant digits and rounds half away from zero, as Excel does
                return (decimal)number;

            if (type == CellValueType.DateTime && TryParseDate(value, out var date))
                return (decimal)date.ToOADate();

            return null;
        }

        private static string FormatNumber(Spec spec, decimal number, bool signed, CultureInfo culture)
        {
            if (spec.Kind == Kind.Percentage)
                number *= 100;

            var pattern = new StringBuilder();
            pattern.Append(spec.Grouping ? "#," : "");
            pattern.Append(spec.MinIntegerDigits == 0 ? "#" : new string('0', spec.MinIntegerDigits));
            if (spec.Decimals > 0)
                pattern.Append('.').Append('0', spec.MinDecimals).Append('#', spec.Decimals - spec.MinDecimals);

            var digits = Math.Abs(number).ToString(pattern.ToString(), culture.NumberFormat);
            // An optional-only fraction can leave a trailing separator ("1." for "0.##" and 1), which Excel shows; keep it
            if (spec.Decimals > 0 && spec.MinDecimals == 0 && !digits.Contains(culture.NumberFormat.NumberDecimalSeparator))
                digits += culture.NumberFormat.NumberDecimalSeparator;

            var negative = signed && number < 0 && Math.Round(Math.Abs(number), spec.Decimals, MidpointRounding.AwayFromZero) != 0;
            return (negative ? "-" : "") + spec.Prefix + digits + spec.Suffix + (spec.Kind == Kind.Percentage ? "%" : "");
        }

        private static string FormatDate(Spec spec, DateTime date, CultureInfo culture)
        {
            var names = culture.DateTimeFormat;
            var twelveHour = spec.DateParts.Any(p => p.Type == "ampm");
            var sb = new StringBuilder();

            foreach (var (type, length, text) in spec.DateParts)
            {
                sb.Append(type switch
                {
                    "text" => text,
                    "y" => length <= 2 ? (date.Year % 100).ToString("00", culture) : date.Year.ToString("0000", culture),
                    "m" => length switch
                    {
                        1 => date.Month.ToString(culture),
                        2 => date.Month.ToString("00", culture),
                        3 => names.GetAbbreviatedMonthName(date.Month),
                        4 => names.GetMonthName(date.Month),
                        _ => names.GetMonthName(date.Month)[..1]
                    },
                    "d" => length switch
                    {
                        1 => date.Day.ToString(culture),
                        2 => date.Day.ToString("00", culture),
                        3 => names.GetAbbreviatedDayName(date.DayOfWeek),
                        _ => names.GetDayName(date.DayOfWeek)
                    },
                    "h" => Pad(twelveHour ? (date.Hour % 12 == 0 ? 12 : date.Hour % 12) : date.Hour, length, culture),
                    "min" => Pad(date.Minute, length, culture),
                    "s" => Pad(date.Second, length, culture),
                    "ampm" => date.Hour < 12 ? Designator(names.AMDesignator, "AM") : Designator(names.PMDesignator, "PM"),
                    _ => ""
                });
            }
            return sb.ToString();
        }

        private static string Pad(int value, int length, CultureInfo culture) =>
            value.ToString(length >= 2 ? "00" : "0", culture);

        private static string Designator(string designator, string fallback) =>
            string.IsNullOrEmpty(designator) ? fallback : designator;
    }
}
