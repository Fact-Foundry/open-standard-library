using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace OslSpreadsheet.Services
{
    /// <summary>
    /// Handles XML serialization and deserialization for ODS file components.
    /// </summary>
    internal static class XmlService
    {
        /// <summary>
        /// Parses XML from a package entry with DTDs prohibited and no external resolver, so entity-expansion
        /// and external-entity attacks are rejected. Neither ODF nor OOXML permits DTDs.
        /// </summary>
        internal static XDocument LoadSecure(Stream stream, bool preserveWhitespace = false)
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = false
            };
            using var reader = XmlReader.Create(stream, settings);
            return XDocument.Load(reader, preserveWhitespace ? LoadOptions.PreserveWhitespace : LoadOptions.None);
        }

        /// <summary>
        /// Throws if the entries in a package declare more uncompressed bytes than the limit allows.
        /// .NET stops reading an entry at its declared size, so the declared total is what bounds memory.
        /// </summary>
        internal static void EnforceSizeLimit(ZipArchive archive, OoxSpreadsheet.ImportOptions options)
        {
            if (!DeclaredSizeWithinLimit(archive.Entries.Select(e => e.Length), options.MaxUncompressedBytes, out long declared))
                throw new OoxSpreadsheet.ImportLimitExceededException(
                    $"The file declares at least {declared:N0} bytes of uncompressed content, which exceeds the import limit of {options.MaxUncompressedBytes:N0} bytes (ImportOptions.MaxUncompressedBytes).");
        }

        /// <summary>
        /// Checks the declared uncompressed sizes of a package's entries against a limit. Each entry is checked on its own
        /// and the total is summed with overflow detection, because the sizes come from the file: Zip64 lets an entry
        /// declare up to 2^63 bytes, and three such entries would wrap an unchecked sum to a negative number that
        /// passes a plain comparison. <paramref name="declared"/> is the total, or the size that failed the check.
        /// </summary>
        internal static bool DeclaredSizeWithinLimit(IEnumerable<long> entryLengths, long limit, out long declared)
        {
            declared = 0;
            foreach (var length in entryLengths)
            {
                if (length < 0 || length > limit)
                {
                    declared = length < 0 ? long.MaxValue : length;
                    return false;
                }

                try
                {
                    declared = checked(declared + length);
                }
                catch (OverflowException)
                {
                    declared = long.MaxValue;
                    return false;
                }

                if (declared > limit)
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Serializes an object to UTF-8 encoded XML bytes. Serialization errors propagate to the caller so a failure
        /// is never written into a package as an empty part.
        /// </summary>
        internal static async Task<byte[]> ConvertToXmlAsync(object obj)
        {
            XmlSerializer serializer = new(obj.GetType());

            var settings = new XmlWriterSettings
            {
                Indent = false,
                OmitXmlDeclaration = false,
                Encoding = new UTF8Encoding(false),
                DoNotEscapeUriAttributes = true
            };

            var ns = new XmlSerializerNamespaces();
            ns.Add(string.Empty, string.Empty);

            await using MemoryStream memoryStream = new();
            using (var writer = XmlWriter.Create(memoryStream, settings))
                serializer.Serialize(writer, obj, ns);

            return memoryStream.ToArray();
        }
    }
}
