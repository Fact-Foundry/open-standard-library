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
            long total = 0;
            foreach (var entry in archive.Entries)
                total += entry.Length;

            if (total > options.MaxUncompressedBytes)
                throw new OoxSpreadsheet.ImportLimitExceededException(
                    $"The file declares {total:N0} bytes of uncompressed content, which exceeds the import limit of {options.MaxUncompressedBytes:N0} bytes (ImportOptions.MaxUncompressedBytes).");
        }

        /// <summary>
        /// Serializes an object to UTF-8 encoded XML bytes.
        /// </summary>
        internal static async Task<byte[]> ConvertToXmlAsync(object obj)
        {
            try
            {
                XmlSerializer serializer = new(obj.GetType());

                var utf8NoBom = new UTF8Encoding(false);
                var settings = new XmlWriterSettings
                {
                    Indent = false,
                    OmitXmlDeclaration = false,
                    Encoding = utf8NoBom,
                    DoNotEscapeUriAttributes = true
                };

                var ns = new XmlSerializerNamespaces();
                ns.Add(string.Empty, string.Empty);

                await using MemoryStream memoryStream = new();
                using var streamWriter = XmlWriter.Create(memoryStream, settings);

                serializer.Serialize(streamWriter, obj, ns);

                var file = memoryStream.ToArray();

                var result = Encoding.UTF8.GetString(file);

                return Encoding.UTF8.GetBytes(result.Replace("utf-8", "UTF-8").Replace("\" />", "\"/>"));
            }
            catch
            {
                return [];
            }
        }
    }
}
