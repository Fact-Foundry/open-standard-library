using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Read-only view of a ZIP-based package (XLSX or ODS) that reports packaging and XML well-formedness problems
    /// to an <see cref="IssueCollector"/> instead of throwing.
    /// </summary>
    internal sealed class ZipPackage : IDisposable
    {
        private readonly MemoryStream _stream;
        private readonly ZipArchive _archive;
        private readonly IssueCollector _issues;
        private readonly string _codePrefix;
        private readonly Dictionary<string, ZipArchiveEntry> _entries = new(StringComparer.Ordinal);
        private readonly Dictionary<string, XDocument?> _xmlCache = new(StringComparer.Ordinal);

        private ZipPackage(MemoryStream stream, ZipArchive archive, IssueCollector issues, string codePrefix)
        {
            _stream = stream;
            _archive = archive;
            _issues = issues;
            _codePrefix = codePrefix;
        }

        /// <summary>
        /// Names of all file entries (directory entries excluded), in archive order.
        /// </summary>
        internal List<string> PartNames { get; } = new();

        /// <summary>
        /// Opens the package and checks entry names, duplicates, and declared size. Returns null if the archive can't be read.
        /// </summary>
        internal static ZipPackage? Open(byte[] file, IssueCollector issues, ValidationOptions options, string codePrefix)
        {
            var stream = new MemoryStream(file, writable: false);
            ZipArchive archive;
            try
            {
                archive = new ZipArchive(stream, ZipArchiveMode.Read);
                _ = archive.Entries.Count;
            }
            catch (InvalidDataException ex)
            {
                stream.Dispose();
                issues.Error($"{codePrefix}_ZIP_CORRUPT", $"The file is not a readable ZIP archive: {ex.Message} The file may be truncated or was not written as binary.");
                return null;
            }

            var package = new ZipPackage(stream, archive, issues, codePrefix);
            if (!package.IndexEntries(options))
            {
                package.Dispose();
                return null;
            }

            return package;
        }

        private bool IndexEntries(ValidationOptions options)
        {
            long totalSize = 0;

            foreach (var entry in _archive.Entries)
            {
                totalSize += entry.Length;
                var name = entry.FullName;

                if (name.Contains('\\'))
                    _issues.Error($"{_codePrefix}_ZIP_BACKSLASH_PATH", $"ZIP entry \"{name}\" uses backslashes. Entry paths must use forward slashes (/).", name);

                if (name.StartsWith('/'))
                    _issues.Error($"{_codePrefix}_ZIP_ABSOLUTE_PATH", $"ZIP entry \"{name}\" starts with a slash. Entry paths must be relative to the archive root.", name);

                if (name.EndsWith('/'))
                    continue;

                if (!_entries.TryAdd(name, entry))
                {
                    _issues.Error($"{_codePrefix}_ZIP_DUPLICATE_ENTRY", $"The archive contains more than one entry named \"{name}\".", name);
                    continue;
                }

                PartNames.Add(name);
            }

            if (totalSize > options.MaxUncompressedBytes)
            {
                _issues.Error($"{_codePrefix}_ZIP_TOO_LARGE",
                    $"The archive declares {totalSize:N0} bytes of uncompressed content, which exceeds the limit of {options.MaxUncompressedBytes:N0} bytes.");
                return false;
            }

            return true;
        }

        internal bool Exists(string partName) => _entries.ContainsKey(partName);

        /// <summary>
        /// Finds an entry whose name matches <paramref name="partName"/> ignoring case. Used to explain near-miss paths.
        /// </summary>
        internal string? FindCaseInsensitive(string partName) =>
            PartNames.FirstOrDefault(n => string.Equals(n, partName, StringComparison.OrdinalIgnoreCase));

        /// <summary>
        /// Reads the raw bytes of an entry, or reports an error and returns null if it can't be decompressed.
        /// </summary>
        internal byte[]? ReadBytes(string partName)
        {
            if (!_entries.TryGetValue(partName, out var entry))
                return null;

            try
            {
                using var entryStream = entry.Open();
                using var ms = new MemoryStream();
                entryStream.CopyTo(ms);
                return ms.ToArray();
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or NotSupportedException)
            {
                _issues.Error($"{_codePrefix}_ZIP_ENTRY_UNREADABLE", $"The entry could not be decompressed: {ex.Message}", partName);
                return null;
            }
        }

        /// <summary>
        /// Parses an entry as XML. Reports a well-formedness error once and returns null if the XML is invalid.
        /// </summary>
        internal XDocument? LoadXml(string partName)
        {
            if (_xmlCache.TryGetValue(partName, out var cached))
                return cached;

            XDocument? doc = null;
            var bytes = ReadBytes(partName);

            // Applications write some placeholder parts as zero-byte files (e.g. LibreOffice's Configurations2/accelerator/current.xml).
            if (bytes != null && bytes.Length > 0)
            {
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    CloseInput = true
                };

                try
                {
                    using var reader = XmlReader.Create(new MemoryStream(bytes), settings);
                    doc = XDocument.Load(reader, LoadOptions.SetLineInfo);
                }
                catch (XmlException ex)
                {
                    _issues.Error($"{_codePrefix}_XML_MALFORMED",
                        $"The XML is not well-formed: {ex.Message} Check for unescaped '&' or '<' characters, mismatched tags, and undeclared namespace prefixes.",
                        partName, $"line {ex.LineNumber}, position {ex.LinePosition}");
                }
            }

            _xmlCache[partName] = doc;
            return doc;
        }

        /// <summary>
        /// Formats the source line of an XML node for issue locations, e.g. "line 12".
        /// </summary>
        internal static string? LineOf(XObject node) =>
            node is IXmlLineInfo info && info.HasLineInfo() ? $"line {info.LineNumber}" : null;

        /// <summary>
        /// Combines a description of a position with the XML line number, e.g. "cell B3, line 12".
        /// </summary>
        internal static string At(string description, XObject node)
        {
            var line = LineOf(node);
            return line == null ? description : $"{description}, {line}";
        }

        public void Dispose()
        {
            _archive.Dispose();
            _stream.Dispose();
        }
    }
}
