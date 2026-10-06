namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Checks whether files are well-formed XLSX, ODS, or delimited text files without importing them.
    /// Use it to reject malformed files (for example, files produced by an LLM) before handing them to a user,
    /// and pass <see cref="ValidationResult.ToString"/> or <see cref="ValidationResult.Issues"/> back to the producer so it can fix them.
    /// </summary>
    /// <remarks>
    /// Validation is structural: it checks the packaging, XML, references, and cell values that spreadsheet applications
    /// depend on to open a file without errors or repair prompts. It is not a full schema validation.
    /// </remarks>
    public static class SpreadsheetValidator
    {
        /// <summary>
        /// Validates a file as the format implied by its file name extension
        /// (.xlsx/.xlsm/.xltx/.xltm, .ods, or .csv/.tsv/.tab/.psv/.txt).
        /// If the extension is not recognized, the format is detected from the file's content.
        /// </summary>
        public static ValidationResult Validate(byte[] file, string fileName, ValidationOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(fileName);
            options ??= new ValidationOptions();

            var extension = Path.GetExtension(fileName).ToLowerInvariant();
            var format = extension switch
            {
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => ValidationFileFormat.Xlsx,
                ".ods" or ".ots" => ValidationFileFormat.Ods,
                ".csv" or ".tsv" or ".tab" or ".psv" or ".txt" => ValidationFileFormat.Delimited,
                _ => ValidationFileFormat.Unknown
            };

            if (format == ValidationFileFormat.Delimited && options.Delimiter == null)
            {
                var inferred = extension switch
                {
                    ".tsv" or ".tab" => ColumnDelimeter.Tab,
                    ".psv" => ColumnDelimeter.Pipe,
                    _ => ColumnDelimeter.Comma
                };
                return ValidateDelimited(file, inferred, options);
            }

            return Validate(file, format, options);
        }

        /// <summary>
        /// Validates a file as the given format. Pass <see cref="ValidationFileFormat.Unknown"/> to detect the format from the content.
        /// </summary>
        public static ValidationResult Validate(byte[] file, ValidationFileFormat format, ValidationOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(file);
            options ??= new ValidationOptions();

            if (format == ValidationFileFormat.Unknown)
                format = DetectFormat(file);

            return format switch
            {
                ValidationFileFormat.Xlsx => ValidateXlsx(file, options),
                ValidationFileFormat.Ods => ValidateOds(file, options),
                ValidationFileFormat.Delimited => ValidateDelimited(file, options),
                _ => Run(ValidationFileFormat.Unknown, options, issues =>
                    issues.Error("FORMAT_UNRECOGNIZED", $"The file is not a recognized spreadsheet format; it appears to be {FileSniffer.Describe(file)}."))
            };
        }

        /// <summary>
        /// Validates an Office Open XML spreadsheet (.xlsx, .xlsm).
        /// </summary>
        public static ValidationResult ValidateXlsx(byte[] file, ValidationOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(file);
            options ??= new ValidationOptions();
            return Run(ValidationFileFormat.Xlsx, options, issues => XlsxValidator.Validate(file, issues, options));
        }

        /// <summary>
        /// Validates an OpenDocument spreadsheet (.ods).
        /// </summary>
        public static ValidationResult ValidateOds(byte[] file, ValidationOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(file);
            options ??= new ValidationOptions();
            return Run(ValidationFileFormat.Ods, options, issues => OdsValidator.Validate(file, issues, options));
        }

        /// <summary>
        /// Validates a delimited text file using <see cref="ValidationOptions.Delimiter"/> (comma when not set)
        /// and <see cref="ValidationOptions.Encoding"/>.
        /// </summary>
        public static ValidationResult ValidateDelimited(byte[] file, ValidationOptions? options = null)
        {
            options ??= new ValidationOptions();
            return ValidateDelimited(file, options.Delimiter ?? ColumnDelimeter.Comma, options);
        }

        /// <summary>
        /// Validates a delimited text file using the given delimiter.
        /// </summary>
        public static ValidationResult ValidateDelimited(byte[] file, ColumnDelimeter delimiter, ValidationOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(file);
            options ??= new ValidationOptions();
            return Run(ValidationFileFormat.Delimited, options, issues => DelimitedValidator.Validate(file, issues, options, delimiter));
        }

        /// <summary>
        /// Determines a file's spreadsheet format from its content. Returns <see cref="ValidationFileFormat.Unknown"/>
        /// for empty files and for content that is not a spreadsheet (PDF, HTML, JSON, legacy .xls, other ZIP files, ...).
        /// </summary>
        public static ValidationFileFormat DetectFormat(byte[] file)
        {
            ArgumentNullException.ThrowIfNull(file);
            return FileSniffer.DetectFormat(file);
        }

        private static ValidationResult Run(ValidationFileFormat format, ValidationOptions options, Action<IssueCollector> validate)
        {
            var issues = new IssueCollector(options);
            try
            {
                validate(issues);
            }
            catch (IssueLimitReachedException)
            {
                // The collector already recorded that the result is truncated.
            }
            return issues.ToResult(format);
        }
    }
}
