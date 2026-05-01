using OslSpreadsheet.Models;
using System.IO.Compression;

namespace OslSpreadsheet.Services
{
    public static class ZipService
    {
        internal static async Task<byte[]> GenerateZipAsync(List<InMemoryFile> files)
        {
            await using MemoryStream archiveStream = new();
            using (ZipArchive archive = new(archiveStream, ZipArchiveMode.Create, true))
            {
                foreach (var file in files)
                {
                    var level = file.Store ? CompressionLevel.NoCompression : CompressionLevel.Fastest;
                    var zipArchiveEntry = archive.CreateEntry(file.FileName, level);
                    using var zipStream = zipArchiveEntry.Open();
                    zipStream.Write(file.Content, 0, file.Content.Length);
                }
            }

            byte[] data = archiveStream.ToArray();

            await archiveStream.DisposeAsync();

            return data;
        }

        internal static byte[] GenerateZip(List<InMemoryFile> files)
        {
            using MemoryStream archiveStream = new();

            using (ZipArchive archive = new(archiveStream, ZipArchiveMode.Create, true))
            {
                foreach (var file in files)
                {
                    var level = file.Store ? CompressionLevel.NoCompression : CompressionLevel.Fastest;
                    var zipArchiveEntry = archive.CreateEntry(file.FileName, level);
                    using var zipStream = zipArchiveEntry.Open();
                    zipStream.Write(file.Content, 0, file.Content.Length);
                }
            }

            byte[] data = archiveStream.ToArray();

            archiveStream.Dispose();

            return data;
        }
    }
}
