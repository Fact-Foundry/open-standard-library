using System.Buffers.Binary;
using OoxSpreadsheet;
using OslSpreadsheet.Services;
using OslSpreadsheet.Validation;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// The import size limit is enforced on sizes the file declares, so it must hold up against sizes chosen by an attacker:
/// single entries larger than the limit, and totals that would wrap a plain sum around to a negative number.
/// </summary>
public class DeclaredSizeTests
{
    private const long Limit = 64L * 1024 * 1024;

    [Fact]
    public void WithinLimit_Passes()
    {
        Assert.True(XmlService.DeclaredSizeWithinLimit(new long[] { 1000, 2000, Limit - 3000 }, Limit, out var total));
        Assert.Equal(Limit, total);
    }

    [Fact]
    public void Empty_Passes()
    {
        Assert.True(XmlService.DeclaredSizeWithinLimit(Array.Empty<long>(), Limit, out var total));
        Assert.Equal(0, total);
    }

    [Fact]
    public void TotalOneOverLimit_Fails()
    {
        Assert.False(XmlService.DeclaredSizeWithinLimit(new long[] { Limit, 1 }, Limit, out var total));
        Assert.Equal(Limit + 1, total);
    }

    [Fact]
    public void SingleEntryOverLimit_Fails()
    {
        Assert.False(XmlService.DeclaredSizeWithinLimit(new long[] { 10, Limit + 1 }, Limit, out var declared));
        Assert.Equal(Limit + 1, declared);
    }

    /// <summary>
    /// Three entries of 2^62 sum to 2^63 + 2^62, which wraps to a negative number in unchecked arithmetic and would pass "total > limit".
    /// </summary>
    [Fact]
    public void SumThatWrapsNegative_Fails()
    {
        var huge = 1L << 62;
        Assert.True(unchecked(huge + huge + huge) < 0, "test premise: the plain sum wraps negative");

        Assert.False(XmlService.DeclaredSizeWithinLimit(new[] { huge, huge, huge }, Limit, out _));
        Assert.False(XmlService.DeclaredSizeWithinLimit(new[] { huge, huge, huge }, long.MaxValue, out _));
    }

    [Fact]
    public void NegativeDeclaredSize_Fails()
    {
        Assert.False(XmlService.DeclaredSizeWithinLimit(new long[] { 10, -1 }, Limit, out _));
    }

    // ----- Real packages with lying size fields -----

    /// <summary>
    /// Rewrites the uncompressed-size field of every entry in a zip (local headers and central directory) to the given
    /// 32-bit value, so the archive stays readable but declares sizes that are not true.
    /// </summary>
    private static byte[] DeclareUncompressedSize(byte[] zip, uint declared)
    {
        var bytes = (byte[])zip.Clone();
        for (int i = 0; i + 30 <= bytes.Length; i++)
        {
            var signature = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i));
            if (signature == 0x04034b50) // local file header: uncompressed size at offset 22
                BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(i + 22), declared);
            else if (signature == 0x02014b50) // central directory header: uncompressed size at offset 24
                BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(i + 24), declared);
        }
        return bytes;
    }

    private static async Task<(byte[] xlsx, byte[] ods)> SmallPackages()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Data").AddCell(1, 1, "x");
        return (await spreadsheet.GenerateXlsxFileAsync(), await spreadsheet.GenerateOdsFileAsync());
    }

    [Fact]
    public async Task Import_RejectsEntryDeclaringSizeOverLimit()
    {
        var (xlsx, ods) = await SmallPackages();
        var options = new ImportOptions { MaxUncompressedBytes = 1_000_000 };

        using var spreadsheet = new Spreadsheet();
        var ex1 = await Assert.ThrowsAsync<ImportLimitExceededException>(() => spreadsheet.ImportXlsxFileAsync(DeclareUncompressedSize(xlsx, 0xFFFF_FFF0), options));
        var ex2 = await Assert.ThrowsAsync<ImportLimitExceededException>(() => spreadsheet.ImportOdsFileAsync(DeclareUncompressedSize(ods, 0xFFFF_FFF0), options));

        Assert.Contains("MaxUncompressedBytes", ex1.Message);
        Assert.Contains("MaxUncompressedBytes", ex2.Message);
    }

    [Fact]
    public async Task Validator_ReportsEntryDeclaringSizeOverLimit()
    {
        var (xlsx, ods) = await SmallPackages();
        var options = new ValidationOptions { MaxUncompressedBytes = 1_000_000 };

        var xlsxResult = SpreadsheetValidator.ValidateXlsx(DeclareUncompressedSize(xlsx, 0xFFFF_FFF0), options);
        var odsResult = SpreadsheetValidator.ValidateOds(DeclareUncompressedSize(ods, 0xFFFF_FFF0), options);

        Assert.Contains(xlsxResult.Issues, i => i.Code == "XLSX_ZIP_TOO_LARGE");
        Assert.Contains(odsResult.Issues, i => i.Code == "ODS_ZIP_TOO_LARGE");
    }

    [Fact]
    public async Task UnmodifiedPackages_StillImport()
    {
        var (xlsx, ods) = await SmallPackages();

        using var spreadsheet = new Spreadsheet();
        Assert.Equal("x", (await spreadsheet.ImportXlsxFileAsync(xlsx)).Sheets[0].GetCell(1, 1)?.Value);
        Assert.Equal("x", (await spreadsheet.ImportOdsFileAsync(ods)).Sheets[0].GetCell(1, 1)?.Value);
    }
}
