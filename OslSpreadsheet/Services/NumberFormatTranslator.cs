using System.Security;
using System.Text;
using System.Xml.Linq;

namespace OslSpreadsheet.Services
{
    /// <summary>
    /// Translates between Excel number format codes (used by <c>CellStyle.NumberFormat</c> and XLSX) and
    /// OpenDocument data styles (the <c>number:*-style</c> elements used by ODS).
    /// </summary>
    /// <remarks>
    /// Supported: fixed decimals, thousands grouping, percentages, currency symbols, text literals, and date/time
    /// patterns built from y, m, d, h, s, and AM/PM. Not supported: multiple sections (positive;negative), colors,
    /// conditions, scientific notation, fractions, and fill/padding characters. Unsupported codes are still written
    /// to XLSX as-is, since Excel understands them, but can't be translated to ODS.
    /// </remarks>
    internal static class NumberFormatTranslator
    {
        internal const string DefaultDateCode = "yyyy-mm-dd";
        internal const string DefaultDateTimeCode = "yyyy-mm-dd hh:mm:ss";

        internal enum Kind { Number, Percentage, Currency, Date, Time, Text, Unsupported }

        /// <summary>
        /// The parsed form of an Excel format code.
        /// </summary>
        internal sealed record Spec(
            Kind Kind,
            int Decimals,
            int MinIntegerDigits,
            bool Grouping,
            string Prefix,
            string Suffix,
            IReadOnlyList<(string Type, int Length, string Text)> DateParts)
        {
            /// <summary>
            /// The currency symbol, when the format is a currency: whichever of prefix or suffix carries text.
            /// </summary>
            internal string CurrencySymbol => Prefix.Trim().Length > 0 ? Prefix : Suffix;

            internal bool CurrencySymbolFirst => Prefix.Trim().Length > 0;
        }

        /// <summary>
        /// Excel's built-in number formats by numFmtId. Files only store the id for these, not the code.
        /// </summary>
        internal static readonly IReadOnlyDictionary<int, string> BuiltInFormats = new Dictionary<int, string>
        {
            [0] = "General", [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00",
            [9] = "0%", [10] = "0.00%", [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??",
            [14] = "m/d/yyyy", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy",
            [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm", [21] = "h:mm:ss", [22] = "m/d/yyyy h:mm",
            [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)", [39] = "#,##0.00;(#,##0.00)", [40] = "#,##0.00;[Red](#,##0.00)",
            [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@"
        };

        /// <summary>
        /// Returns the built-in numFmtId for a format code, or null if it needs a custom numFmt entry.
        /// </summary>
        internal static int? BuiltInId(string code)
        {
            foreach (var (id, builtIn) in BuiltInFormats)
                if (builtIn == code)
                    return id;
            return null;
        }

        // ---------------------------------------------------------------- Excel code parsing

        /// <summary>
        /// Parses an Excel format code. Only the first section is used.
        /// </summary>
        internal static Spec Parse(string code)
        {
            var section = SplitSections(code)[0];
            var parts = new List<(string Type, int Length, string Text)>();   // date parts: y, M (month), m (minute), d, D (weekday), h, s, ampm, text
            var literal = new StringBuilder();
            bool percent = false, hasDigits = false, hasText = false, grouping = false, seenDot = false, unsupported = false;
            int decimals = 0, minInteger = 0;
            var prefix = new StringBuilder();
            var suffix = new StringBuilder();

            void FlushLiteral()
            {
                if (literal.Length == 0) return;
                parts.Add(("text", 0, literal.ToString()));
                literal.Clear();
            }

            int i = 0;
            while (i < section.Length)
            {
                var c = section[i];

                if (c == '"')
                {
                    int end = section.IndexOf('"', i + 1);
                    if (end < 0) end = section.Length;
                    var text = section[(i + 1)..end];
                    literal.Append(text);
                    (hasDigits ? suffix : prefix).Append(text);
                    i = end + 1;
                    continue;
                }
                if (c == '\\' && i + 1 < section.Length)
                {
                    literal.Append(section[i + 1]);
                    (hasDigits ? suffix : prefix).Append(section[i + 1]);
                    i += 2;
                    continue;
                }
                if (c == '[')
                {
                    int end = section.IndexOf(']', i + 1);
                    if (end < 0) end = section.Length;
                    var inner = section[(i + 1)..end];
                    if (inner.StartsWith('$'))
                    {
                        // Locale-tagged currency symbol, e.g. [$€-407]
                        var symbol = inner[1..].Split('-')[0];
                        literal.Append(symbol);
                        (hasDigits ? suffix : prefix).Append(symbol);
                    }
                    else if (inner.Length > 0 && inner.All(ch => "hms".Contains(char.ToLowerInvariant(ch))))
                    {
                        unsupported = true; // elapsed time [h], [mm], [ss]
                    }
                    // Colors and conditions ([Red], [>100]) are dropped.
                    i = end + 1;
                    continue;
                }
                if (c == '_' && i + 1 < section.Length)
                {
                    literal.Append(' ');
                    (hasDigits ? suffix : prefix).Append(' ');
                    i += 2;
                    continue;
                }
                if (c == '*' && i + 1 < section.Length)
                {
                    i += 2; // fill character: no fixed-width equivalent
                    continue;
                }

                if (section.AsSpan(i).StartsWith("AM/PM", StringComparison.OrdinalIgnoreCase))
                {
                    FlushLiteral();
                    parts.Add(("ampm", 0, ""));
                    i += 5;
                    continue;
                }
                if (section.AsSpan(i).StartsWith("A/P", StringComparison.OrdinalIgnoreCase))
                {
                    FlushLiteral();
                    parts.Add(("ampm", 0, ""));
                    i += 3;
                    continue;
                }

                var lower = char.ToLowerInvariant(c);
                if ("ymdhs".Contains(lower))
                {
                    int start = i;
                    while (i < section.Length && char.ToLowerInvariant(section[i]) == lower) i++;
                    FlushLiteral();
                    parts.Add((lower.ToString(), i - start, ""));
                    continue;
                }

                switch (c)
                {
                    case '#' or '0' or '?':
                        hasDigits = true;
                        if (seenDot) decimals++;
                        else if (c == '0') minInteger++;
                        break;
                    case ',':
                        if (hasDigits && i + 1 < section.Length && "#0?".Contains(section[i + 1])) grouping = true;
                        else if (hasDigits) unsupported = true; // trailing comma scales by 1000
                        else literal.Append(c);
                        break;
                    case '.':
                        if (hasDigits && !seenDot) seenDot = true;
                        else literal.Append(c);
                        break;
                    case '%':
                        percent = true;
                        break;
                    case '@':
                        hasText = true;
                        break;
                    case 'E' or 'e' when hasDigits:
                        unsupported = true; // scientific notation
                        i++;
                        break;
                    case '/' when hasDigits:
                        unsupported = true; // fraction
                        break;
                    default:
                        literal.Append(c);
                        (hasDigits ? suffix : prefix).Append(c);
                        break;
                }
                i++;
            }
            FlushLiteral();

            var dateParts = parts.Where(p => p.Type != "text").ToList();
            var isDate = dateParts.Count > 0;

            if (unsupported || (isDate && (hasDigits || hasText || percent)) || (hasText && hasDigits))
                return new Spec(Kind.Unsupported, 0, 0, false, "", "", parts);

            if (isDate)
            {
                ResolveMinutes(parts);
                var timeOnly = parts.All(p => p.Type is "h" or "min" or "s" or "ampm" or "text");
                return new Spec(timeOnly ? Kind.Time : Kind.Date, 0, 0, false, "", "", parts);
            }

            if (hasText)
                return new Spec(Kind.Text, 0, 0, false, "", "", parts);

            if (!hasDigits)
                return new Spec(Kind.Unsupported, 0, 0, false, "", "", parts);

            var p = prefix.ToString();
            var s = suffix.ToString();
            var kind = percent ? Kind.Percentage : (p.Trim().Length > 0 || s.Trim().Length > 0) ? Kind.Currency : Kind.Number;
            return new Spec(kind, decimals, minInteger, grouping, p, s, parts);
        }

        /// <summary>
        /// An 'm' run means minutes when it follows hours or precedes seconds; otherwise it is a month.
        /// </summary>
        private static void ResolveMinutes(List<(string Type, int Length, string Text)> parts)
        {
            var dateIdx = parts.Select((p, i) => (p, i)).Where(x => x.p.Type != "text").ToList();
            for (int k = 0; k < dateIdx.Count; k++)
            {
                if (dateIdx[k].p.Type != "m") continue;
                var prev = k > 0 ? dateIdx[k - 1].p.Type : null;
                var next = k + 1 < dateIdx.Count ? dateIdx[k + 1].p.Type : null;
                if (prev == "h" || next == "s")
                    parts[dateIdx[k].i] = ("min", dateIdx[k].p.Length, "");
            }
        }

        private static List<string> SplitSections(string code)
        {
            var sections = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            foreach (var c in code)
            {
                if (c == '"') inQuotes = !inQuotes;
                if (c == ';' && !inQuotes) { sections.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(c);
            }
            sections.Add(sb.ToString());
            return sections;
        }

        // ---------------------------------------------------------------- Excel -> ODS

        /// <summary>
        /// Builds the ODS data style element for a format code, or null if the code can't be translated.
        /// </summary>
        internal static string? ToOdsDataStyle(string code, string styleName)
        {
            var spec = Parse(code);
            var sb = new StringBuilder();

            switch (spec.Kind)
            {
                case Kind.Number:
                    sb.Append($"<number:number-style style:name=\"{styleName}\">");
                    AppendText(sb, spec.Prefix);
                    AppendNumber(sb, spec);
                    AppendText(sb, spec.Suffix);
                    sb.Append("</number:number-style>");
                    break;

                case Kind.Percentage:
                    sb.Append($"<number:percentage-style style:name=\"{styleName}\">");
                    AppendText(sb, spec.Prefix);
                    AppendNumber(sb, spec);
                    AppendText(sb, spec.Suffix + "%");
                    sb.Append("</number:percentage-style>");
                    break;

                case Kind.Currency:
                    sb.Append($"<number:currency-style style:name=\"{styleName}\">");
                    if (spec.CurrencySymbolFirst)
                    {
                        sb.Append($"<number:currency-symbol>{SecurityElement.Escape(spec.Prefix)}</number:currency-symbol>");
                        AppendNumber(sb, spec);
                        AppendText(sb, spec.Suffix);
                    }
                    else
                    {
                        AppendText(sb, spec.Prefix);
                        AppendNumber(sb, spec);
                        sb.Append($"<number:currency-symbol>{SecurityElement.Escape(spec.Suffix)}</number:currency-symbol>");
                    }
                    sb.Append("</number:currency-style>");
                    break;

                case Kind.Date:
                case Kind.Time:
                    var element = spec.Kind == Kind.Time ? "number:time-style" : "number:date-style";
                    sb.Append($"<{element} style:name=\"{styleName}\">");
                    foreach (var (type, length, text) in spec.DateParts)
                    {
                        string Style(int longAt) => length >= longAt ? "long" : "short";
                        sb.Append(type switch
                        {
                            "text" => $"<number:text>{SecurityElement.Escape(text)}</number:text>",
                            "y" => $"<number:year number:style=\"{Style(4)}\"/>",
                            "m" when length >= 3 => $"<number:month number:style=\"{Style(4)}\" number:textual=\"true\"/>",
                            "m" => $"<number:month number:style=\"{Style(2)}\"/>",
                            "d" when length >= 3 => $"<number:day-of-week number:style=\"{Style(4)}\"/>",
                            "d" => $"<number:day number:style=\"{Style(2)}\"/>",
                            "h" => $"<number:hours number:style=\"{Style(2)}\"/>",
                            "min" => $"<number:minutes number:style=\"{Style(2)}\"/>",
                            "s" => $"<number:seconds number:style=\"{Style(2)}\"/>",
                            "ampm" => "<number:am-pm/>",
                            _ => ""
                        });
                    }
                    sb.Append($"</{element}>");
                    break;

                case Kind.Text:
                    sb.Append($"<number:text-style style:name=\"{styleName}\"><number:text-content/></number:text-style>");
                    break;

                default:
                    return null;
            }

            return sb.ToString();
        }

        private static void AppendNumber(StringBuilder sb, Spec spec)
        {
            sb.Append($"<number:number number:decimal-places=\"{spec.Decimals}\" number:min-decimal-places=\"{spec.Decimals}\" number:min-integer-digits=\"{spec.MinIntegerDigits}\"");
            if (spec.Grouping) sb.Append(" number:grouping=\"true\"");
            sb.Append("/>");
        }

        private static void AppendText(StringBuilder sb, string text)
        {
            if (text.Length > 0)
                sb.Append($"<number:text>{SecurityElement.Escape(text)}</number:text>");
        }

        /// <summary>
        /// The ODS value type that matches a format code: "percentage", "currency", or null for a plain number.
        /// </summary>
        internal static string? OdsValueType(string code) => Parse(code).Kind switch
        {
            Kind.Percentage => "percentage",
            Kind.Currency => "currency",
            _ => null
        };

        /// <summary>
        /// ISO 4217 code for a currency symbol, for the office:currency attribute. Null when unknown.
        /// </summary>
        internal static string? CurrencyCode(string symbol)
        {
            var s = symbol.Trim();
            if (s.Length == 3 && s.All(char.IsAsciiLetterUpper)) return s;
            return s switch { "$" => "USD", "€" => "EUR", "£" => "GBP", "¥" => "JPY", _ => null };
        }

        // ---------------------------------------------------------------- ODS -> Excel

        private static readonly XNamespace NumberNs = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

        /// <summary>
        /// Converts an ODS data style element to an Excel format code, or null if it has no equivalent.
        /// </summary>
        internal static string? FromOdsDataStyle(XElement style)
        {
            var sb = new StringBuilder();
            var local = style.Name.LocalName;

            if (local == "text-style") return "@";
            if (local is not ("number-style" or "percentage-style" or "currency-style" or "date-style" or "time-style")) return null;

            foreach (var child in style.Elements())
            {
                if (child.Name.Namespace != NumberNs) continue;
                string Len(string longValue, string shortValue) => (string?)child.Attribute(NumberNs + "style") == "long" ? longValue : shortValue;

                switch (child.Name.LocalName)
                {
                    case "text":
                        var text = child.Value;
                        if (local == "percentage-style" && text == "%") { sb.Append('%'); break; }
                        sb.Append(text.All(ch => "/-:. ,()".Contains(ch)) ? text : $"\"{text.Replace("\"", "")}\"");
                        break;
                    case "currency-symbol":
                        sb.Append($"\"{child.Value.Replace("\"", "")}\"");
                        break;
                    case "number":
                        // A bare <number:number> with no decimal-places (LibreOffice's N0 / "General") has no fixed-format equivalent
                        if (child.Attribute(NumberNs + "decimal-places") == null && style.Elements().Count() == 1)
                            return null;
                        int decimals = (int?)child.Attribute(NumberNs + "decimal-places") ?? 0;
                        int minInt = (int?)child.Attribute(NumberNs + "min-integer-digits") ?? 1;
                        bool grouping = (string?)child.Attribute(NumberNs + "grouping") == "true";
                        var integer = (grouping, minInt) switch
                        {
                            (true, 0) => "#,###",
                            (true, _) => "#,##" + new string('0', minInt),
                            (false, 0) => "#",
                            (false, _) => new string('0', minInt)
                        };
                        sb.Append(integer);
                        if (decimals > 0) sb.Append('.').Append('0', decimals);
                        break;
                    case "scientific-number":
                        int sciDecimals = (int?)child.Attribute(NumberNs + "decimal-places") ?? 2;
                        sb.Append("0.").Append('0', sciDecimals).Append("E+00");
                        break;
                    case "fraction":
                        return null;
                    case "year": sb.Append(Len("yyyy", "yy")); break;
                    case "month":
                        bool textual = (string?)child.Attribute(NumberNs + "textual") == "true";
                        sb.Append(textual ? Len("mmmm", "mmm") : Len("mm", "m"));
                        break;
                    case "day": sb.Append(Len("dd", "d")); break;
                    case "day-of-week": sb.Append(Len("dddd", "ddd")); break;
                    case "hours": sb.Append(Len("hh", "h")); break;
                    case "minutes": sb.Append(Len("mm", "m")); break;
                    case "seconds": sb.Append(Len("ss", "s")); break;
                    case "am-pm": sb.Append("AM/PM"); break;
                    case "era" or "quarter" or "week-of-year": return null;
                }
            }

            return sb.Length > 0 ? sb.ToString() : null;
        }
    }
}
