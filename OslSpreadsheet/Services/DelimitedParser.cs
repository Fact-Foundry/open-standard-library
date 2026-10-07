using System.Runtime.CompilerServices;
using System.Text;

namespace OslSpreadsheet.Services
{
    /// <summary>
    /// Reads records from delimited text. Comma, tab, and pipe files use RFC 4180 quoting: a value may be wrapped in
    /// double quotes, inside which delimiters and line breaks are data and a quote is written as two quotes.
    /// ASCII-delimited files use the unit separator (0x1F) between values and the record separator (0x1E) between rows, with no quoting.
    /// </summary>
    /// <remarks>
    /// Parsing is lenient, like spreadsheet applications: a quote inside an unquoted value is kept as data, text after a
    /// closing quote is appended to the value, and an unterminated quoted value runs to the end of the input.
    /// Use <c>SpreadsheetValidator</c> to detect those problems. Blank lines are skipped.
    /// </remarks>
    internal static class DelimitedParser
    {
        private const char UnitSeparator = '\u001F';
        private const char RecordSeparator = '\u001E';

        internal static char GetDelimiter(ColumnDelimeter delimiter) => delimiter switch
        {
            ColumnDelimeter.ASCII => UnitSeparator,
            ColumnDelimeter.Pipe => '|',
            ColumnDelimeter.Tab => '\t',
            _ => ','
        };

        internal static async IAsyncEnumerable<string[]> ParseAsync(TextReader reader, ColumnDelimeter delimiter,
            [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            var separator = GetDelimiter(delimiter);
            var isAscii = delimiter == ColumnDelimeter.ASCII;

            var fields = new List<string>();
            var field = new StringBuilder();
            var inQuotes = false;
            var afterClosingQuote = false;
            var recordHasContent = false;
            var pendingCarriageReturn = false;
            var firstChar = true;

            var buffer = new char[16 * 1024];
            int read;
            while ((read = await reader.ReadAsync(buffer.AsMemory(), cancellationToken)) > 0)
            {
                for (int i = 0; i < read; i++)
                {
                    var c = buffer[i];

                    if (firstChar)
                    {
                        firstChar = false;
                        if (c == '﻿')
                            continue;
                    }

                    // A CR LF pair ends a single record.
                    if (pendingCarriageReturn)
                    {
                        pendingCarriageReturn = false;
                        if (c == '\n')
                            continue;
                    }

                    if (inQuotes)
                    {
                        if (c == '"')
                        {
                            inQuotes = false;
                            afterClosingQuote = true;
                        }
                        else
                            field.Append(c);
                        continue;
                    }

                    if (afterClosingQuote && c == '"')
                    {
                        // Doubled quote inside a quoted value
                        field.Append('"');
                        inQuotes = true;
                        afterClosingQuote = false;
                        continue;
                    }

                    if (c == separator)
                    {
                        fields.Add(field.ToString());
                        field.Clear();
                        afterClosingQuote = false;
                        recordHasContent = true;
                        continue;
                    }

                    var isRecordEnd = isAscii ? c == RecordSeparator : c is '\n' or '\r';
                    if (isRecordEnd)
                    {
                        pendingCarriageReturn = c == '\r' && !isAscii;
                        if (recordHasContent || field.Length > 0 || afterClosingQuote)
                        {
                            fields.Add(field.ToString());
                            yield return fields.ToArray();
                        }
                        fields.Clear();
                        field.Clear();
                        afterClosingQuote = false;
                        recordHasContent = false;
                        continue;
                    }

                    if (c == '"' && !isAscii && field.Length == 0 && !afterClosingQuote)
                    {
                        inQuotes = true;
                        recordHasContent = true;
                        continue;
                    }

                    afterClosingQuote = false;
                    field.Append(c);
                }
            }

            if (recordHasContent || field.Length > 0 || afterClosingQuote || inQuotes)
            {
                fields.Add(field.ToString());
                yield return fields.ToArray();
            }
        }

        /// <summary>
        /// Formats a value for a delimited file. Comma files always quote values; tab and pipe files quote only values
        /// that would otherwise be misread (containing the delimiter or a line break, or starting with a quote).
        /// ASCII-delimited values are written as-is.
        /// </summary>
        internal static string FormatValue(string value, ColumnDelimeter delimiter)
        {
            if (delimiter == ColumnDelimeter.ASCII)
                return value;

            var needsQuotes = delimiter == ColumnDelimeter.Comma
                || value.IndexOf(GetDelimiter(delimiter)) >= 0
                || value.IndexOfAny(['\r', '\n']) >= 0
                || value.StartsWith('"');

            return needsQuotes ? $"\"{value.Replace("\"", "\"\"")}\"" : value;
        }
    }
}
