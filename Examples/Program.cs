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

await TestRoundTripOds();
await TestRoundTripXlsx();
await TestAllFeatures();
await TestOdsMinimal();

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

async Task TestRoundTripOds()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;
        var sheet1 = await workbook.AddSheetAsync("Data");
        sheet1.AddCell(1, 1, "Name");
        sheet1.AddCell(1, 2, "Score");
        sheet1.AddCell(2, 1, "Alice");
        sheet1.AddCell(2, 2, "95.5", CellValueType.Float);
        sheet1.AddCell(3, 1, "Bob");
        sheet1.AddCell(3, 2, "82.0", CellValueType.Float);

        var odsFile = await spreadsheet.GenerateOdsFileAsync();
        var imported = await spreadsheet.ImportOdsFileAsync(odsFile);

        Console.WriteLine("=== ODS Round-trip ===");
        foreach (var sheet in imported.Sheets)
        {
            Console.WriteLine($"Sheet: {sheet.SheetName}");
            foreach (var cell in sheet.Cells)
                Console.WriteLine($"  [{cell.Row},{cell.Column}] ({cell.ValueType}) = {cell.Value}");
        }
    }
}

async Task TestRoundTripXlsx()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;
        var sheet1 = await workbook.AddSheetAsync("Data");
        sheet1.AddCell(1, 1, "Name");
        sheet1.AddCell(1, 2, "Score");
        sheet1.AddCell(2, 1, "Alice");
        sheet1.AddCell(2, 2, "95.5", CellValueType.Float);
        sheet1.AddCell(3, 1, "Bob");
        sheet1.AddCell(3, 2, "82.0", CellValueType.Float);

        var xlsxFile = await spreadsheet.GenerateXlsxFileAsync();
        var imported = await spreadsheet.ImportXlsxFileAsync(xlsxFile);

        Console.WriteLine("=== XLSX Round-trip ===");
        foreach (var sheet in imported.Sheets)
        {
            Console.WriteLine($"Sheet: {sheet.SheetName}");
            foreach (var cell in sheet.Cells)
                Console.WriteLine($"  [{cell.Row},{cell.Column}] ({cell.ValueType}) = {cell.Value}");
        }
    }
}

async Task TestAllFeatures()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;
        workbook.Creator = "Kevin Williams";

        var sheet = await workbook.AddSheetAsync("Data Dictionary");
        sheet.FreezeRows = 1;

        var headerStyle = new CellStyle
        {
            Bold = true,
            FontColor = "#FFFFFF",
            BackgroundColor = "#2196F3",
            BorderBottom = new CellBorder { Style = BorderStyle.Medium, Color = "#1565C0" },
            WrapText = true
        };

        var headers = new[] { "Name", "Type", "Description", "Expression" };
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = sheet.AddCell(1, c + 1, headers[c]);
            cell.Style = headerStyle;
        }

        var zebraStyle = new CellStyle { BackgroundColor = "#F5F5F5" };

        string[] names = { "customer_id", "first_name", "last_name", "email_address", "account_balance", "is_active", "signup_date", "last_login", "total_orders" };
        string[] types = { "Integer", "String", "String", "String", "Decimal", "Boolean", "DateTime", "DateTime", "Integer" };
        string[] descs = {
            "Unique customer identifier",
            "Customer's first name",
            "Customer's last name",
            "Primary email address for communications and login",
            "Current account balance in USD",
            "Whether the customer account is currently active",
            "Date the customer first registered",
            "Most recent login timestamp",
            "Total number of completed orders"
        };

        for (int r = 0; r < names.Length; r++)
        {
            int row = r + 2;
            sheet.AddCell(row, 1, names[r]);
            sheet.AddCell(row, 2, types[r]);
            sheet.AddCell(row, 3, descs[r]);
            sheet.AddCell(row, 4, $"=Fields[\"{names[r]}\"]");

            if (row % 2 == 0)
                foreach (var c in sheet.GetRow(row))
                    c.Style = zebraStyle;
        }

        sheet.AutoFitColumns(minWidth: 10, maxWidth: 60);
        sheet.SetAutoFilter();

        var xlsxFile = await spreadsheet.GenerateXlsxFileAsync();
        await File.WriteAllBytesAsync("/tmp/test_all_features.xlsx", xlsxFile);
        Console.WriteLine("Wrote /tmp/test_all_features.xlsx");

        var odsFile = await spreadsheet.GenerateOdsFileAsync();
        await File.WriteAllBytesAsync("/tmp/test_all_features.ods", odsFile);
        Console.WriteLine("Wrote /tmp/test_all_features.ods");
    }
}

async Task TestOdsMinimal()
{
    // 1: plain data
    await using (var s1 = host.Services.GetService<ISpreadsheet>())
    {
        var sheet = s1.Workbook.AddSheet("Sheet1");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Value");
        sheet.AddCell(2, 1, "A");
        sheet.AddCell(2, 2, "1");
        await File.WriteAllBytesAsync("/tmp/ods_1_plain.ods", await s1.GenerateOdsFileAsync());
        Console.WriteLine("Wrote /tmp/ods_1_plain.ods");
    }

    // 2: with styling
    await using (var s2 = host.Services.GetService<ISpreadsheet>())
    {
        var sheet = s2.Workbook.AddSheet("Sheet1");
        var h = sheet.AddCell(1, 1, "Name");
        h.Style = new CellStyle { Bold = true, BackgroundColor = "#2196F3" };
        sheet.AddCell(1, 2, "Value");
        sheet.AddCell(2, 1, "A");
        sheet.AddCell(2, 2, "1");
        await File.WriteAllBytesAsync("/tmp/ods_2_styled.ods", await s2.GenerateOdsFileAsync());
        Console.WriteLine("Wrote /tmp/ods_2_styled.ods");
    }

    // 3: with column widths
    await using (var s3 = host.Services.GetService<ISpreadsheet>())
    {
        var sheet = s3.Workbook.AddSheet("Sheet1");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Value");
        sheet.AddCell(2, 1, "A");
        sheet.AddCell(2, 2, "1");
        sheet.SetColumnWidth(1, 15);
        sheet.SetColumnWidth(2, 20);
        await File.WriteAllBytesAsync("/tmp/ods_3_colwidth.ods", await s3.GenerateOdsFileAsync());
        Console.WriteLine("Wrote /tmp/ods_3_colwidth.ods");
    }

    // 4: with freeze
    await using (var s4 = host.Services.GetService<ISpreadsheet>())
    {
        var sheet = s4.Workbook.AddSheet("Sheet1");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Value");
        sheet.AddCell(2, 1, "A");
        sheet.AddCell(2, 2, "1");
        sheet.FreezeRows = 1;
        await File.WriteAllBytesAsync("/tmp/ods_4_freeze.ods", await s4.GenerateOdsFileAsync());
        Console.WriteLine("Wrote /tmp/ods_4_freeze.ods");
    }

    // 5: with auto-filter only
    await using (var s5 = host.Services.GetService<ISpreadsheet>())
    {
        var sheet = s5.Workbook.AddSheet("Sheet1");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Value");
        sheet.AddCell(2, 1, "A");
        sheet.AddCell(2, 2, "1");
        sheet.SetAutoFilter();
        await File.WriteAllBytesAsync("/tmp/ods_5_filter.ods", await s5.GenerateOdsFileAsync());
        Console.WriteLine("Wrote /tmp/ods_5_filter.ods");
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
        await File.WriteAllBytesAsync(@"/tmp/test.xlsx", xlsxFile);
    }
}
