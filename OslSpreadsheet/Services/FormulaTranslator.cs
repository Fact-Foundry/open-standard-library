using System.Text;

namespace OslSpreadsheet.Services
{
    /// <summary>
    /// Converts formulas between Excel A1 syntax (used by <c>oCell.Formula</c> and XLSX) and OpenFormula syntax (used by ODS).
    /// </summary>
    /// <remarks>
    /// Excel: <c>=SUM(A1:B2, Sheet2!$C$3)</c>. OpenFormula: <c>of:=SUM([.A1:.B2]; [$Sheet2.$C$3])</c>.
    /// Only references, argument separators, and the namespace prefix differ; function names and operators are passed through.
    /// </remarks>
    internal static class FormulaTranslator
    {
        private const string OpenFormulaPrefix = "of:";

        /// <summary>
        /// A cell, column, or row reference part such as <c>$A$1</c>, <c>A</c>, or <c>1</c>.
        /// </summary>
        private sealed record RefPart(bool ColAbsolute, string? Column, bool RowAbsolute, int? Row)
        {
            internal bool IsCell => Column != null && Row != null;

            public override string ToString() =>
                (Column == null ? "" : (ColAbsolute ? "$" : "") + Column) + (Row == null ? "" : (RowAbsolute ? "$" : "") + Row);
        }

        /// <param name="External">True when the reference points into another workbook, e.g. <c>[1]Sheet1!A1</c>.</param>
        private sealed record Reference(string? Sheet, RefPart Start, RefPart? End, bool External);

        /// <summary>
        /// Ensures a formula starts with '=' and returns null for blank formulas.
        /// </summary>
        internal static string? Normalize(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return null;
            var trimmed = formula.Trim();
            return trimmed.StartsWith('=') ? trimmed : "=" + trimmed;
        }

        /// <summary>
        /// Returns the formula text as stored in an XLSX <c>&lt;f&gt;</c> element (no leading '=').
        /// </summary>
        internal static string ToXlsx(string formula) => Normalize(formula)![1..];

        /// <summary>
        /// Converts an XLSX <c>&lt;f&gt;</c> element value to the <c>oCell.Formula</c> form.
        /// </summary>
        internal static string FromXlsx(string formula) => "=" + formula;

        /// <summary>
        /// Converts an Excel A1 formula to an ODS <c>table:formula</c> value, e.g. <c>=SUM(A1:A3)</c> to <c>of:=SUM([.A1:.A3])</c>.
        /// </summary>
        internal static string ToOpenFormula(string formula)
        {
            var body = Normalize(formula)![1..];
            var sb = new StringBuilder(OpenFormulaPrefix).Append('=');

            RewriteExcel(body, sb, text => text.Replace(',', ';'), reference =>
            {
                var sheet = reference.Sheet == null ? "" : "$" + QuoteOpenFormulaSheet(reference.Sheet);
                var result = $"[{sheet}.{reference.Start}";
                if (reference.End != null)
                    result += $":.{reference.End}";
                return result + "]";
            });

            return sb.ToString();
        }

        /// <summary>
        /// Converts an ODS <c>table:formula</c> value to Excel A1 syntax, e.g. <c>of:=SUM([.A1:.A3])</c> to <c>=SUM(A1:A3)</c>.
        /// </summary>
        internal static string FromOpenFormula(string formula)
        {
            // Strip the namespace prefix (of:, oooc:, msoxl:) that precedes the '='.
            var equals = formula.IndexOf('=');
            var prefixEnd = formula.IndexOf(':');
            var body = equals >= 0 && (prefixEnd < 0 || prefixEnd > equals) ? formula[(equals + 1)..]
                : prefixEnd >= 0 ? formula[(prefixEnd + 1)..].TrimStart('=')
                : formula;

            var sb = new StringBuilder("=");
            int i = 0;
            while (i < body.Length)
            {
                var c = body[i];
                if (c == '"')
                {
                    i = CopyStringLiteral(body, i, sb);
                }
                else if (c == '[')
                {
                    var close = FindClosingBracket(body, i);
                    sb.Append(ConvertOpenFormulaReference(body[(i + 1)..close]));
                    i = close + 1;
                }
                else
                {
                    sb.Append(c == ';' ? ',' : c);
                    i++;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Shifts the relative parts of every reference in an Excel formula. Used to expand XLSX shared formulas,
        /// which store the formula once and apply it to other cells with relative references adjusted.
        /// </summary>
        internal static string Shift(string formula, int rowOffset, int columnOffset)
        {
            var sb = new StringBuilder();
            RewriteExcel(formula, sb, text => text, reference =>
            {
                var sheet = reference.Sheet == null ? "" : QuoteExcelSheet(reference.Sheet) + "!";
                var result = sheet + ShiftPart(reference.Start, rowOffset, columnOffset);
                if (reference.End != null)
                    result += ":" + ShiftPart(reference.End, rowOffset, columnOffset);
                return result;
            });
            return sb.ToString();
        }

        /// <summary>
        /// Returns the sheet names referenced by an Excel formula (e.g. "My Data" for <c>='My Data'!A1</c>),
        /// excluding references into other workbooks.
        /// </summary>
        internal static IReadOnlyCollection<string> ReferencedSheets(string formula)
        {
            var sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            RewriteExcel(Normalize(formula) ?? "", new StringBuilder(), text => text, reference =>
            {
                if (reference.Sheet != null && !reference.External)
                    sheets.Add(reference.Sheet);
                return "";
            });
            return sheets;
        }

        private static RefPart ShiftPart(RefPart part, int rowOffset, int columnOffset) => part with
        {
            Column = part.Column == null || part.ColAbsolute ? part.Column : ColumnName(ColumnNumber(part.Column) + columnOffset),
            Row = part.Row == null || part.RowAbsolute ? part.Row : part.Row + rowOffset
        };

        // ---------------------------------------------------------------- Excel tokenizer

        /// <summary>
        /// Walks an Excel formula, passing literal text through <paramref name="mapText"/> and references through
        /// <paramref name="mapReference"/>. String literals are copied unchanged.
        /// </summary>
        private static void RewriteExcel(string formula, StringBuilder sb, Func<string, string> mapText, Func<Reference, string> mapReference)
        {
            var text = new StringBuilder();
            void FlushText()
            {
                if (text.Length == 0) return;
                sb.Append(mapText(text.ToString()));
                text.Clear();
            }

            int i = 0;
            while (i < formula.Length)
            {
                var c = formula[i];

                if (c == '"')
                {
                    FlushText();
                    i = CopyStringLiteral(formula, i, sb);
                    continue;
                }

                var atBoundary = i == 0 || !IsIdentifierChar(formula[i - 1]);
                if (atBoundary && (c == '\'' || c == '$' || char.IsAsciiLetterOrDigit(c) || c == '_'))
                {
                    if (TryParseReference(formula, i, out var reference, out var end))
                    {
                        FlushText();
                        sb.Append(mapReference(reference));
                        i = end;
                        continue;
                    }

                    // Not a reference: consume the whole identifier or number so references aren't matched inside it.
                    var start = i;
                    if (c == '\'')
                        i++;
                    while (i < formula.Length && (IsIdentifierChar(formula[i]) || formula[i] == '$'))
                        i++;
                    text.Append(formula, start, Math.Max(1, i - start));
                    i = Math.Max(i, start + 1);
                    continue;
                }

                text.Append(c);
                i++;
            }
            FlushText();
        }

        private static bool TryParseReference(string s, int start, out Reference reference, out int end)
        {
            reference = null!;
            end = start;
            int i = start;
            string? sheet = null;

            // Optional sheet prefix: 'My Sheet'! or Sheet1!
            if (s[i] == '\'')
            {
                var name = new StringBuilder();
                i++;
                while (i < s.Length)
                {
                    if (s[i] == '\'' && i + 1 < s.Length && s[i + 1] == '\'') { name.Append('\''); i += 2; continue; }
                    if (s[i] == '\'') break;
                    name.Append(s[i++]);
                }
                if (i + 1 >= s.Length || s[i] != '\'' || s[i + 1] != '!')
                    return false;
                sheet = name.ToString();
                i += 2;
            }
            else
            {
                int j = i;
                while (j < s.Length && IsIdentifierChar(s[j])) j++;
                if (j < s.Length && j > i && s[j] == '!')
                {
                    sheet = s[i..j];
                    i = j + 1;
                }
            }

            var first = ParsePart(s, ref i);
            if (first == null)
                return false;

            RefPart? second = null;
            if (i < s.Length && s[i] == ':')
            {
                int j = i + 1;
                second = ParsePart(s, ref j);
                if (second != null && (second.Column == null) == (first.Column == null) && (second.Row == null) == (first.Row == null))
                    i = j;
                else
                    second = null;
            }

            // Column-only (A:A) and row-only (1:1) parts are only references as part of a range.
            if (!first.IsCell && second == null)
                return false;

            // Anything identifier-like directly after means this was a name or function (e.g. LOG10( or Rate2024x).
            if (i < s.Length && (IsIdentifierChar(s[i]) || s[i] == '(' || s[i] == '!'))
                return false;

            reference = new Reference(sheet, first, second, External: start > 0 && s[start - 1] == ']');
            end = i;
            return true;
        }

        private static RefPart? ParsePart(string s, ref int i)
        {
            int j = i;
            bool colAbs = false, rowAbs = false;
            string? column = null;
            int? row = null;

            if (j < s.Length && s[j] == '$') { colAbs = true; j++; }
            int letters = j;
            while (j < s.Length && char.IsAsciiLetter(s[j])) j++;
            if (j > letters)
            {
                if (j - letters > 3) return null;
                column = s[letters..j].ToUpperInvariant();
            }
            else if (colAbs)
            {
                // "$1" is an absolute row, not an absolute column.
                colAbs = false;
                rowAbs = true;
            }

            if (j < s.Length && s[j] == '$' && column != null) { rowAbs = true; j++; }
            int digits = j;
            while (j < s.Length && char.IsAsciiDigit(s[j])) j++;
            if (j > digits)
            {
                if (j - digits > 7 || s[digits] == '0') return null;
                row = int.Parse(s.AsSpan(digits, j - digits));
            }
            else if (rowAbs && column != null)
                return null;

            if (column == null && row == null)
                return null;

            i = j;
            return new RefPart(colAbs, column, rowAbs, row);
        }

        // ---------------------------------------------------------------- OpenFormula references

        private static string ConvertOpenFormulaReference(string inner)
        {
            // inner: "$Sheet1.A1:.B2", ".A1", "'My Sheet'.A1", ".A:.A", or error forms such as ".#REF!"
            var parts = SplitOutsideQuotes(inner, ':');
            string? sheet = null;
            var refs = new List<string>();

            foreach (var part in parts)
            {
                var p = part.TrimStart('$');
                var dot = LastDotOutsideQuotes(p);
                if (dot < 0)
                    return inner.Contains("#REF!") ? "#REF!" : inner;

                var sheetPart = p[..dot];
                if (sheetPart.Length > 0 && sheet == null)
                    sheet = UnquoteOpenFormulaSheet(sheetPart);
                refs.Add(p[(dot + 1)..]);
            }

            if (refs.Any(r => r.Contains("#REF!")))
                return "#REF!";

            var prefix = sheet == null ? "" : QuoteExcelSheet(sheet) + "!";
            return prefix + string.Join(':', refs);
        }

        private static List<string> SplitOutsideQuotes(string s, char separator)
        {
            var parts = new List<string>();
            var inQuotes = false;
            int start = 0;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '\'') inQuotes = !inQuotes;
                else if (s[i] == separator && !inQuotes)
                {
                    parts.Add(s[start..i]);
                    start = i + 1;
                }
            }
            parts.Add(s[start..]);
            return parts;
        }

        private static int LastDotOutsideQuotes(string s)
        {
            var inQuotes = false;
            var last = -1;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '\'') inQuotes = !inQuotes;
                else if (s[i] == '.' && !inQuotes) last = i;
            }
            return last;
        }

        private static int FindClosingBracket(string s, int open)
        {
            var inQuotes = false;
            for (int i = open + 1; i < s.Length; i++)
            {
                if (s[i] == '\'') inQuotes = !inQuotes;
                else if (s[i] == ']' && !inQuotes) return i;
            }
            return s.Length - 1;
        }

        // ---------------------------------------------------------------- Helpers

        private static int CopyStringLiteral(string s, int start, StringBuilder sb)
        {
            int i = start + 1;
            while (i < s.Length)
            {
                if (s[i] == '"')
                {
                    if (i + 1 < s.Length && s[i + 1] == '"') { i += 2; continue; }
                    i++;
                    break;
                }
                i++;
            }
            sb.Append(s, start, i - start);
            return i;
        }

        private static bool IsIdentifierChar(char c) => char.IsAsciiLetterOrDigit(c) || c == '_' || c == '.';

        private static string QuoteOpenFormulaSheet(string sheet) =>
            sheet.All(c => char.IsAsciiLetterOrDigit(c) || c == '_') ? sheet : $"'{sheet.Replace("'", "''")}'";

        private static string UnquoteOpenFormulaSheet(string sheet) =>
            sheet.StartsWith('\'') && sheet.EndsWith('\'') && sheet.Length >= 2 ? sheet[1..^1].Replace("''", "'") : sheet;

        private static string QuoteExcelSheet(string sheet)
        {
            var plain = sheet.Length > 0 && (char.IsAsciiLetter(sheet[0]) || sheet[0] == '_')
                && sheet.All(c => char.IsAsciiLetterOrDigit(c) || c == '_' || c == '.')
                && !LooksLikeCellReference(sheet);
            return plain ? sheet : $"'{sheet.Replace("'", "''")}'";
        }

        private static bool LooksLikeCellReference(string name)
        {
            int i = 0;
            while (i < name.Length && char.IsAsciiLetter(name[i])) i++;
            return i is > 0 and <= 3 && i < name.Length && name[i..].All(char.IsAsciiDigit);
        }

        private static int ColumnNumber(string column)
        {
            int n = 0;
            foreach (var c in column)
                n = n * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
            return n;
        }

        private static string ColumnName(int column)
        {
            var name = "";
            while (column > 0)
            {
                column--;
                name = (char)('A' + column % 26) + name;
                column /= 26;
            }
            return name;
        }
    }
}
