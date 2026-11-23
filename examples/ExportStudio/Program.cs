using DataExportPackage;
using ExportFiles.Services;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapGet("/api/health", () => Results.Ok(new { status = "ok", packages = new[] { "DataExport 2.0.0", "ExcelExportKit 2.0.0" } }));

app.MapGet("/api/preview", () => Results.Ok(SampleData.Orders));

app.MapGet("/api/export/{mode}/{format}", async (string mode, string format, CancellationToken cancellationToken) =>
{
    mode = mode.ToLowerInvariant();
    format = format.ToLowerInvariant();
    if (!(mode == "data" && format is "csv" or "json" or "xml" or "xlsx") &&
        !(mode == "workbook" && format is "csv" or "json" or "pdf" or "docx") &&
        !(mode == "report" && format == "xlsx"))
        return Results.BadRequest(new { error = "Unsupported export mode or format." });

    var folder = Path.Combine(Path.GetTempPath(), "exportstudio-" + Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(folder);
    try
    {
        cancellationToken.ThrowIfCancellationRequested();
        var output = Path.Combine(folder, "export." + format);
        if (mode == "data")
        {
            var exporter = DataExporter.FromObjects(SampleData.Orders);
            switch (format)
            {
                case "csv":
                    exporter.ToCsv(output, new DataExportOptions { CsvProtectFormulas = true });
                    break;
                case "json":
                    exporter.ToJson(output);
                    break;
                case "xml":
                    exporter.ToXml(output);
                    break;
                case "xlsx":
                    exporter.ToExcel(output, new DataExportOptions
                    {
                        WorksheetName = "Orders", ExcelIncludeHeaders = true,
                        ExcelFreezeHeader = true, ExcelAutoFilter = true
                    });
                    break;
            }
        }
        else if (mode == "report")
        {
            using var report = new ExcelReportBuilder();
            report.AddSheet("Sales", SampleData.Orders,
                new ExcelReportSheetOptions { AutoFitColumns = true });
            report.AddTotals("Sales", 5, 8);
            report.AddSheet("Regions", SampleData.Regions);
            report.SaveAs(output);
        }
        else
        {
            // Demo pipeline: C# objects -> DataExport XLSX -> ExcelExportKit conversion.
            var workbook = Path.Combine(folder, "source.xlsx");
            DataExporter.FromObjects(SampleData.Orders).ToExcel(workbook,
                new DataExportOptions { WorksheetName = "Orders", ExcelIncludeHeaders = true });
            var options = new ExcelConversionOptions
            {
                WorksheetName = "Orders",
                DocumentTitle = "Export Studio — Sales Orders",
                PdfLandscape = true,
                PdfPreserveCellFormatting = true,
                PreserveJsonTypes = true,
                CsvProtectFormulas = true,
                MaxWorkbookFileBytes = 4 * 1024 * 1024,
                MaxUncompressedWorkbookBytes = 16 * 1024 * 1024,
                MaxWorkbookArchiveEntries = 256,
                MaxSelectedCells = 2000
            };
            switch (format)
            {
                case "csv": await ExcelConverter.ToCSVAsync(workbook, output, options); break;
                case "json": await ExcelConverter.ToJsonAsync(workbook, output, options); break;
                case "pdf": await ExcelConverter.ToPDFAsync(workbook, output, options); break;
                case "docx": await ExcelConverter.ToWordAsync(workbook, output, options); break;
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        var bytes = await File.ReadAllBytesAsync(output, cancellationToken);
        var mime = format switch
        {
            "csv" => "text/csv; charset=utf-8",
            "json" => "application/json",
            "xml" => "application/xml",
            "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "pdf" => "application/pdf",
            "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            _ => "application/octet-stream"
        };
        return Results.File(bytes, mime, $"export-studio-{mode}.{format}");
    }
    finally
    {
        Directory.Delete(folder, recursive: true);
    }
});

app.Run();

internal static class SampleData
{
    internal sealed record Order(string OrderId, string Customer, string Region, string Product,
        int Units, decimal UnitPrice, DateTime OrderedAt, decimal Revenue);

    internal sealed record RegionSummary(string Region, int Orders, decimal TotalRevenue);

    internal static readonly Order[] Orders =
    [
        new("SO-1042", "Amina Karimova", "Tashkent", "Analytics license", 3, 149m, new(2026, 11, 2), 447m),
        new("SO-1043", "Jasur Rahimov", "Samarkand", "Team plan", 5, 89m, new(2026, 11, 4), 445m),
        new("SO-1044", "Malika Rustamova", "Bukhara", "Export toolkit", 2, 199m, new(2026, 11, 5), 398m),
        new("SO-1045", "Aziz Yuldashev", "Tashkent", "Analytics license", 4, 149m, new(2026, 11, 8), 596m),
        new("SO-1046", "Dilnoza Akbarova", "Fergana", "Team plan", 8, 89m, new(2026, 11, 10), 712m),
        new("SO-1047", "Behruz Tursunov", "Namangan", "Export toolkit", 3, 199m, new(2026, 11, 12), 597m),
        new("SO-1048", "Sabina Ismoilova", "Tashkent", "Team plan", 6, 89m, new(2026, 11, 15), 534m),
        new("SO-1049", "Otabek Sodiqov", "Andijan", "Analytics license", 2, 149m, new(2026, 11, 17), 298m)
    ];

    internal static readonly RegionSummary[] Regions = Orders
        .GroupBy(x => x.Region)
        .Select(g => new RegionSummary(g.Key, g.Count(), g.Sum(x => x.Revenue)))
        .OrderByDescending(x => x.TotalRevenue)
        .ToArray();
}
