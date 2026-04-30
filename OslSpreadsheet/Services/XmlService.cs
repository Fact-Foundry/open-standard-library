using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace OslSpreadsheet.Services
{
    internal static class XmlService
    {
        internal static async Task<byte[]> ConvertToXmlAsync(object obj)
        {
            try
            {
                // Generate serializer
                XmlSerializer serializer = new(obj.GetType());

                // Settings
                var settings = new XmlWriterSettings
                {
                    Indent = false,
                    OmitXmlDeclaration = false,
                    Encoding = Encoding.UTF8,
                    DoNotEscapeUriAttributes = true
                };

                var ns = new XmlSerializerNamespaces();
                ns.Add(string.Empty, string.Empty);

                // Generate memory stream
                await using MemoryStream memoryStream = new();
                using var streamWriter = XmlWriter.Create(memoryStream, settings);
                streamWriter.WriteStartDocument(true);

                serializer.Serialize(streamWriter, obj, ns);

                var file = memoryStream.ToArray();

                var result = Encoding.UTF8.GetString(file);

                return Encoding.UTF8.GetBytes(result.Replace("utf-8", "UTF-8").Replace("\" />", "\"/>"));

                //return file;
            }
            catch (Exception ex)
            {
                return new byte[0];
            }
        }

        //async Task TestConvertXmlToObject()
        //{
        //    var xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        //        + "<document-meta xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\" >"
        //        + "<meta>"
        //        + "<generator>Open Standard Library v1.0.0</generator>"
        //        + "<initial-creator>Open Standard Library v1.0.0</initial-creator>"
        //        + "<creator>Open Standard Library v1.0.0</creator>"
        //        + "<creation-date>2022-01-08T23:01:46Z </creation-date>"
        //        + "<date>2022-01-08T23:01:46Z</date>"
        //        + "</meta>"
        //        + "</document-meta>";

        //    var obj = await XmlService.ConvertToObject<ODMeta>(xml);
        //}

        /// <summary>
        /// Converts XML to a model
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="xml"></param>
        /// <returns></returns>
        internal static async Task<object> ConvertToObject<T>(string xml)
        {
            T retval = default;

            using (var reader = new StringReader(xml))
            {
                retval = (T)new XmlSerializer(typeof(T)).Deserialize(reader);
            }

            return retval;
        }
    }
}
