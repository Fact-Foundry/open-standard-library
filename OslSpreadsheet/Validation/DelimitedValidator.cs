using System.Text;

namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Validates delimited text files. Comma-delimited files follow RFC 4180 quoting rules; tab- and pipe-delimited
    /// files allow (but don't require) RFC 4180 quoting; ASCII-delimited files use unit (0x1F) and record (0x1E)
    /// separators with no quoting.
    /// </summary>
    internal sealed class DelimitedValidator
    {
        private const string Prefix = "CSV";

        private readonly IssueCollector _issues;
        private readonly ValidationOptions _options;
        private readonly char _delimiter;
        private readonly bool _isAsciiDelimited;
        private readonly bool _strictQuoting;

        private DelimitedValidator(IssueCollector issues, ValidationOptions options, ColumnDelimeter delimiter)
        {
            _issues = issues;
            _options = options;
            _isAsciiDelimited = delimiter == ColumnDelimeter.ASCII;
            _strictQuoting = delimiter == ColumnDelimeter.Comma;
            _delimiter = delimiter switch
            {
                ColumnDelimeter.ASCII => '\u001F',
                ColumnDelimeter.Pipe => '|',
                ColumnDelimeter.Tab => '\t',
                _ => ','
            };
        }

        internal static void Validate(byte[] file, IssueCollector issues, ValidationOptions options, ColumnDelimeter delimiter)
        {
            var validator = new DelimitedValidator(issues, options, delimiter);
            var text = validator.Decode(file);
            if (text != null)
                validator.Parse(text);
        }

        private string? Decode(byte[] file)
        {
            if (file.Length == 0)
            {
                _issues.Error($"{Prefix}_EMPTY", "The file is empty.");
                return null;
            }

            var description = FileSniffer.Describe(file);
            if (FileSniffer.IsZip(file) || description.StartsWith("a legacy", StringComparison.Ordinal) || description == "a PDF document"
                || description == "unrecognized binary data")
            {
                _issues.Error($"{Prefix}_NOT_TEXT", $"A delimited file must be plain text, but this file is {description}.");
                return null;
            }

            var hasUtf16Bom = file.Length >= 2 && ((file[0] == 0xFF && file[1] == 0xFE) || (file[0] == 0xFE && file[1] == 0xFF));
            if (_options.Encoding is FileEncoding.UTF8 or FileEncoding.ASCII && hasUtf16Bom)
            {
                _issues.Error($"{Prefix}_ENCODING_MISMATCH",
                    $"The file starts with a UTF-16 byte order mark but {_options.Encoding} was expected. Save it as {_options.Encoding} or validate with Encoding = Unicode.");
                return null;
            }

            Encoding encoding = _options.Encoding switch
            {
                FileEncoding.ASCII => Encoding.GetEncoding("us-ascii", EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback),
                FileEncoding.Unicode => new UnicodeEncoding(bigEndian: false, byteOrderMark: true, throwOnInvalidBytes: true),
                FileEncoding.UTF32 => new UTF32Encoding(bigEndian: false, byteOrderMark: true, throwOnInvalidCharacters: true),
                _ => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true)
            };

            try
            {
                var text = encoding.GetString(file);
                return text.Length > 0 && text[0] == '﻿' ? text[1..] : text;
            }
            catch (DecoderFallbackException ex)
            {
                _issues.Error($"{Prefix}_ENCODING_INVALID",
                    $"The file is not valid {_options.Encoding} text (invalid byte sequence at byte offset {ex.Index}). Save the file as {_options.Encoding}.");
                return null;
            }
        }

        private enum State { FieldStart, Unquoted, Quoted, QuoteInQuoted }

        private void Parse(string text)
        {
            if (text.Contains('\0'))
            {
                _issues.Error($"{Prefix}_NOT_TEXT", "The file contains NUL characters, so it is binary data rather than delimited text.");
                return;
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                _issues.Error($"{Prefix}_EMPTY", "The file contains no data.");
                return;
            }

            var state = State.FieldStart;
            var line = 1;
            var recordLine = 1;
            var quoteStartLine = 1;
            var fieldLength = 0;
            var fieldCount = 0;
            var recordNumber = 0;
            int? expectedFields = null;
            var allSingleField = true;
            var pendingBlankLines = new List<int>();
            var bareQuoteReportedOnLine = 0;
            var recordLineIsEmpty = true;

            void EndField()
            {
                if (fieldLength > XlsxValidator.MaxCellTextLength)
                    _issues.Warning($"{Prefix}_FIELD_TOO_LONG",
                        $"A field is {fieldLength:N0} characters long. Excel and LibreOffice truncate cell text at {XlsxValidator.MaxCellTextLength:N0} characters.",
                        location: $"row {recordNumber + 1}, line {recordLine}", row: recordNumber + 1);
                fieldCount++;
                fieldLength = 0;
                state = State.FieldStart;
            }

            void EndRecord(bool atEndOfFile)
            {
                EndField();

                // A line with nothing on it is a blank line, not a record with one empty field.
                if (fieldCount == 1 && recordLineIsEmpty)
                {
                    if (!atEndOfFile)
                        pendingBlankLines.Add(recordLine);
                    fieldCount = 0;
                    return;
                }

                foreach (var blank in pendingBlankLines)
                    _issues.Warning($"{Prefix}_BLANK_LINE", "Blank line between rows. Spreadsheet applications import it as an empty row.", location: $"line {blank}");
                pendingBlankLines.Clear();

                recordNumber++;
                if (fieldCount != 1)
                    allSingleField = false;

                if (expectedFields == null)
                    expectedFields = fieldCount;
                else if (fieldCount != expectedFields)
                {
                    var advice = _isAsciiDelimited
                        ? "Values must not contain the unit separator character (0x1F)."
                        : _delimiter == ','
                            ? "Values that contain a comma, double quote, or line break must be wrapped in double quotes, and quotes inside them doubled (\"\")."
                            : $"Values must not contain the {DelimiterName} character unless they are wrapped in double quotes.";
                    var message = $"Row {recordNumber} has {fieldCount} field(s) but the first row has {expectedFields}. {advice}";

                    if (_options.RequireConsistentColumnCount)
                        _issues.Error($"{Prefix}_FIELD_COUNT_MISMATCH", message, location: $"row {recordNumber}, line {recordLine}", row: recordNumber);
                    else
                        _issues.Warning($"{Prefix}_FIELD_COUNT_MISMATCH", message, location: $"row {recordNumber}, line {recordLine}", row: recordNumber);
                }

                fieldCount = 0;
            }

            var quotingEnabled = !_isAsciiDelimited;

            for (int i = 0; i < text.Length; i++)
            {
                var c = text[i];
                var isRecordEnd = _isAsciiDelimited ? c == '\u001E' : c is '\n' or '\r';

                if (state == State.Quoted)
                {
                    if (c == '"')
                        state = State.QuoteInQuoted;
                    else
                    {
                        fieldLength++;
                        if (c == '\n' || (c == '\r' && (i + 1 >= text.Length || text[i + 1] != '\n')))
                            line++;
                    }
                    continue;
                }

                if (state == State.QuoteInQuoted && c == '"')
                {
                    fieldLength++;
                    state = State.Quoted;
                    continue;
                }

                if (c == _delimiter)
                {
                    recordLineIsEmpty = false;
                    EndField();
                    continue;
                }

                if (isRecordEnd)
                {
                    EndRecord(atEndOfFile: false);
                    if (!_isAsciiDelimited)
                    {
                        if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                            i++;
                        line++;
                    }
                    recordLine = line;
                    recordLineIsEmpty = true;
                    continue;
                }

                if (_isAsciiDelimited && c is '\n' or '\r')
                {
                    // Line breaks are ordinary data in ASCII-delimited files but still count for locations.
                    if (c == '\n' || i + 1 >= text.Length || text[i + 1] != '\n')
                        line++;
                }

                recordLineIsEmpty = false;

                switch (state)
                {
                    case State.FieldStart:
                        if (c == '"' && quotingEnabled)
                        {
                            state = State.Quoted;
                            quoteStartLine = line;
                        }
                        else
                        {
                            state = State.Unquoted;
                            fieldLength++;
                            if (c == ' ' && _strictQuoting && NextNonSpaceIsQuote(text, i))
                                _issues.Warning($"{Prefix}_SPACE_BEFORE_QUOTE",
                                    "A field starts with spaces followed by a double quote. The quote is treated as data, not as the start of a quoted value. Remove the spaces after the delimiter.",
                                    location: $"row {recordNumber + 1}, line {line}", row: recordNumber + 1);
                        }
                        break;

                    case State.Unquoted:
                        fieldLength++;
                        if (c == '"' && _strictQuoting && bareQuoteReportedOnLine != line)
                        {
                            bareQuoteReportedOnLine = line;
                            _issues.Warning($"{Prefix}_BARE_QUOTE",
                                "A double quote appears inside a value that is not wrapped in quotes. Wrap the value in double quotes and double the inner quote (\"\").",
                                location: $"row {recordNumber + 1}, line {line}", row: recordNumber + 1);
                        }
                        break;

                    case State.QuoteInQuoted:
                        _issues.Error($"{Prefix}_TEXT_AFTER_CLOSING_QUOTE",
                            $"A quoted value is followed by '{Printable(c)}' instead of a delimiter or line break. A double quote inside a quoted value must be written as two double quotes (\"\").",
                            location: $"row {recordNumber + 1}, line {line}", row: recordNumber + 1);
                        state = State.Unquoted;
                        fieldLength++;
                        break;
                }
            }

            if (state == State.Quoted)
            {
                _issues.Error($"{Prefix}_UNTERMINATED_QUOTE",
                    $"A quoted value that starts on line {quoteStartLine} is never closed, so the rest of the file is read as part of that value. Add the closing double quote, and double any quotes inside the value (\"\").",
                    location: $"row {recordNumber + 1}, line {quoteStartLine}", row: recordNumber + 1);
                return;
            }

            if (!recordLineIsEmpty || state != State.FieldStart || fieldCount > 0)
                EndRecord(atEndOfFile: true);

            if (recordNumber == 0)
            {
                _issues.Error($"{Prefix}_EMPTY", "The file contains no data rows.");
                return;
            }

            if (allSingleField && recordNumber > 1 && DetectOtherDelimiter(text) is string other)
                _issues.Warning($"{Prefix}_POSSIBLE_WRONG_DELIMITER",
                    $"Every row has a single field when split on the {DelimiterName} delimiter, but the text contains {other} characters. The file may use a different delimiter.");
        }

        private string DelimiterName => _delimiter switch
        {
            ',' => "comma",
            '\t' => "tab",
            '|' => "pipe",
            _ => "unit separator"
        };

        private string? DetectOtherDelimiter(string text)
        {
            var candidates = new (char Char, string Name)[] { (',', "comma"), ('\t', "tab"), ('|', "pipe"), (';', "semicolon") };
            var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries).Take(20).ToList();
            foreach (var (ch, name) in candidates)
            {
                if (ch == _delimiter)
                    continue;
                if (lines.Count > 0 && lines.All(l => l.Contains(ch)))
                    return name;
            }
            return null;
        }

        private static bool NextNonSpaceIsQuote(string text, int index)
        {
            while (index < text.Length && text[index] == ' ')
                index++;
            return index < text.Length && text[index] == '"';
        }

        private static string Printable(char c) => char.IsControl(c) ? $"\\u{(int)c:X4}" : c.ToString();
    }
}
