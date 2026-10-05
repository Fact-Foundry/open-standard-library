using System.IO.Compression;
using System.Text;

namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// Identifies what kind of file a byte array actually contains by inspecting its signature and leading content.
    /// </summary>
    internal static class FileSniffer
    {
        internal const string OdsMimeType = "application/vnd.oasis.opendocument.spreadsheet";
        internal const string OdsTemplateMimeType = "application/vnd.oasis.opendocument.spreadsheet-template";

        private static readonly byte[] ZipSignature = [0x50, 0x4B, 0x03, 0x04];
        private static readonly byte[] EmptyZipSignature = [0x50, 0x4B, 0x05, 0x06];
        private static readonly byte[] OleSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        private static readonly byte[] PdfSignature = "%PDF"u8.ToArray();

        internal static bool IsZip(byte[] file) => StartsWith(file, ZipSignature) || StartsWith(file, EmptyZipSignature);

        /// <summary>
        /// Determines the spreadsheet format of a file from its content, ignoring any file name.
        /// </summary>
        internal static ValidationFileFormat DetectFormat(byte[] file)
        {
            if (file.Length == 0)
                return ValidationFileFormat.Unknown;

            if (IsZip(file))
            {
                return DescribeZip(file) switch
                {
                    ZipKind.Ods => ValidationFileFormat.Ods,
                    ZipKind.Xlsx => ValidationFileFormat.Xlsx,
                    _ => ValidationFileFormat.Unknown
                };
            }

            if (StartsWith(file, OleSignature) || StartsWith(file, PdfSignature))
                return ValidationFileFormat.Unknown;

            return DescribeText(file) == TextKind.Plain ? ValidationFileFormat.Delimited : ValidationFileFormat.Unknown;
        }

        /// <summary>
        /// Returns a short noun phrase describing what the file appears to be, e.g. "an XLSX workbook" or "a PDF document".
        /// </summary>
        internal static string Describe(byte[] file)
        {
            if (file.Length == 0)
                return "an empty file";

            if (IsZip(file))
            {
                return DescribeZip(file) switch
                {
                    ZipKind.Ods => "an OpenDocument spreadsheet (ODS)",
                    ZipKind.OtherOpenDocument => "an OpenDocument file that is not a spreadsheet (e.g. a text document or presentation)",
                    ZipKind.Xlsx => "an Office Open XML workbook (XLSX)",
                    ZipKind.Docx => "a Word document (DOCX)",
                    ZipKind.Pptx => "a PowerPoint presentation (PPTX)",
                    ZipKind.Corrupt => "a corrupt or truncated ZIP archive",
                    _ => "a ZIP archive that is not a spreadsheet"
                };
            }

            if (StartsWith(file, OleSignature))
                return "a legacy binary Office file (such as an Excel 97-2003 .xls file)";

            if (StartsWith(file, PdfSignature))
                return "a PDF document";

            return DescribeText(file) switch
            {
                TextKind.Binary => "unrecognized binary data",
                TextKind.Base64Zip => "base64-encoded text of a ZIP-based file (it must be decoded to raw bytes before saving)",
                TextKind.Html => "an HTML document",
                TextKind.SpreadsheetMl2003 => "an Excel 2003 XML spreadsheet (SpreadsheetML), which is not XLSX",
                TextKind.FlatOds => "a flat OpenDocument XML file (FODS), which is not a zipped ODS file",
                TextKind.Xml => "an XML document",
                TextKind.Json => "a JSON document",
                TextKind.Markdown => "Markdown text",
                _ => "plain text (it may be CSV or another delimited format)"
            };
        }

        private enum ZipKind { Ods, OtherOpenDocument, Xlsx, Docx, Pptx, Other, Corrupt }

        private enum TextKind { Plain, Binary, Base64Zip, Html, SpreadsheetMl2003, FlatOds, Xml, Json, Markdown }

        private static ZipKind DescribeZip(byte[] file)
        {
            try
            {
                using var ms = new MemoryStream(file, writable: false);
                using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

                var mimetype = archive.GetEntry("mimetype");
                if (mimetype != null && mimetype.Length < 256)
                {
                    using var reader = new StreamReader(mimetype.Open(), Encoding.ASCII);
                    var value = reader.ReadToEnd().Trim();
                    if (value is OdsMimeType or OdsTemplateMimeType)
                        return ZipKind.Ods;
                    if (value.StartsWith("application/vnd.oasis.opendocument.", StringComparison.Ordinal))
                        return ZipKind.OtherOpenDocument;
                }

                var names = archive.Entries.Select(e => e.FullName.Replace('\\', '/')).ToList();
                if (names.Any(n => n.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)))
                    return ZipKind.Xlsx;
                if (names.Any(n => n.StartsWith("word/", StringComparison.OrdinalIgnoreCase)))
                    return ZipKind.Docx;
                if (names.Any(n => n.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase)))
                    return ZipKind.Pptx;
                if (names.Contains("content.xml") || names.Contains("META-INF/manifest.xml"))
                    return ZipKind.Ods;

                return ZipKind.Other;
            }
            catch (InvalidDataException)
            {
                return ZipKind.Corrupt;
            }
        }

        private static TextKind DescribeText(byte[] file)
        {
            var sampleLength = Math.Min(file.Length, 8192);
            var hasUtf16Bom = file.Length >= 2 && ((file[0] == 0xFF && file[1] == 0xFE) || (file[0] == 0xFE && file[1] == 0xFF));

            if (!hasUtf16Bom && Array.IndexOf(file, (byte)0, 0, sampleLength) >= 0)
                return TextKind.Binary;

            var encoding = hasUtf16Bom ? (file[0] == 0xFF ? Encoding.Unicode : Encoding.BigEndianUnicode) : Encoding.UTF8;
            var text = encoding.GetString(file, 0, sampleLength).TrimStart('﻿', ' ', '\t', '\r', '\n');

            if (text.StartsWith("UEsDB", StringComparison.Ordinal))
                return TextKind.Base64Zip;

            if (text.StartsWith('<'))
            {
                if (text.Contains("<html", StringComparison.OrdinalIgnoreCase) || text.StartsWith("<!DOCTYPE html", StringComparison.OrdinalIgnoreCase))
                    return TextKind.Html;
                if (text.Contains("urn:schemas-microsoft-com:office:spreadsheet", StringComparison.Ordinal))
                    return TextKind.SpreadsheetMl2003;
                if (text.Contains("office:document", StringComparison.Ordinal))
                    return TextKind.FlatOds;
                return TextKind.Xml;
            }

            if (text.StartsWith('{') || (text.StartsWith('[') && text.TrimStart('[', ' ', '\r', '\n', '\t').StartsWith('{')))
                return TextKind.Json;

            if (text.StartsWith("```", StringComparison.Ordinal) || text.Contains("|---", StringComparison.Ordinal) || text.Contains("| ---", StringComparison.Ordinal))
                return TextKind.Markdown;

            return TextKind.Plain;
        }

        private static bool StartsWith(byte[] file, byte[] signature) =>
            file.Length >= signature.Length && file.AsSpan(0, signature.Length).SequenceEqual(signature);
    }
}
