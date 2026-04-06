using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OoxSpreadsheet;
using OslSpreadsheet.Models;

using IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices((_, services) =>
        services
            .AddOptions()
            .AddLogging(configure => configure.AddConsole())
            .Configure<LoggerFilterOptions>(options => options.MinLevel = LogLevel.Information)
            .AddTransient<ISpreadsheet, Spreadsheet>()
    ).Build();

// Begin test code:

await TestGenerateOds();


// Run the application
await host.RunAsync();


async Task TestGenerateCsv()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;

        workbook.Creator = "Kevin Williams";
        workbook.ColumnDelimeter = ColumnDelimeter.Comma;

        var sheet1 = await workbook.AddSheetAsync();

        sheet1.AddCell(1, 1, "Item #");
        sheet1.AddCell(1, 2, "Price");
        sheet1.AddCell(2, 1, "5\" Fitting");
        sheet1.AddCell(2, 2, "10.20", CellValueType.Float);

        // Convert spreadsheet to compressed file (XLSX)
        var csvFile = await spreadsheet.GenerateCsvFileAsync();

        // Save compressed file
        switch(workbook.ColumnDelimeter)
        {
            case ColumnDelimeter.ASCII:
            case ColumnDelimeter.Pipe:
            case ColumnDelimeter.Tab:
                await File.WriteAllBytesAsync(@"C:\Temp\New File.txt", csvFile);
                break;

            case ColumnDelimeter.Comma:
            default:
                await File.WriteAllBytesAsync(@"C:\Temp\New File.csv", csvFile);
                break;
        }
    }
}

async Task TestImportCsv()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var file = File.ReadAllBytes(@"C:\Temp\New File.csv");

        var workbook = await spreadsheet.ImportCsvFileAsync(file);

        var foo = workbook.Sheets.First().ToArray();
    }
}

async Task TestGenerateOds()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;

        workbook.Creator = "Kevin Williams";

        var sheet1 = await workbook.AddSheetAsync();
        var sheet2 = await workbook.AddSheetAsync("Stuff");

        var cell = sheet2.AddCell(1, 1);
        cell.Value = "300.20";
        cell.ValueType = CellValueType.Float;

        var sheet3 = await spreadsheet.Workbook.AddSheetAsync();

        spreadsheet.Workbook.AddSheet("Foo");

        // Convert spreadsheet to compressed file (ODS)
        var odsFile = await spreadsheet.GenerateOdsFileAsync();

        // Save compressed file
        await File.WriteAllBytesAsync(@"/tmp/test.ods", odsFile);
    }
}

async Task TestGenerateXlsx()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;

        workbook.Creator = "Kevin Williams";

        var sheet1 = await workbook.AddSheetAsync();
        var sheet2 = await workbook.AddSheetAsync("Stuff");

        var cell = sheet2.AddCell(1, 1);
        cell.Value = "300.20";
        cell.ValueType = CellValueType.Float;

        var sheet3 = await spreadsheet.Workbook.AddSheetAsync();

        spreadsheet.Workbook.AddSheet("Foo");

        // Convert spreadsheet to compressed file (XLSX)
        var xlsxFile = await spreadsheet.GenerateXlsxFileAsync();

        // Save compressed file
        await File.WriteAllBytesAsync(@"C:\Temp\New File (XLSX).xlsx", xlsxFile);
    }
}
