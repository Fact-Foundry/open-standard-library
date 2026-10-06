using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace OslSpreadsheet.Services
{
    /// <summary>
    /// Handles XML serialization and deserialization for ODS file components.
    /// </summary>
    internal static class XmlService
    {
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

        /// <summary>
        /// Deserializes an XML string into the specified type.
        /// </summary>
        internal static async Task<T?> ConvertToObject<T>(string xml) where T : class
        {
            return await Task.Run(() =>
            {
                using var reader = new StringReader(xml);
                return new XmlSerializer(typeof(T)).Deserialize(reader) as T;
            });
        }
    }
}
