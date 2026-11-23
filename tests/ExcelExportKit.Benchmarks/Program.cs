using System.Diagnostics;
using System.Text.Json;
using ClosedXML.Excel;
using ExportFiles.Services;

if (args.Length is < 1 or > 2 ||
    !int.TryParse(args[0], out var rows) ||
    rows is < 1 or > 100_000)
{
    Console.Error.WriteLine(
        "Usage: dotnet run --project tests/ExcelExportKit.Benchmarks -c Release -- <rows:1..100000> [output-directory]");
    return 2;
}

var outputDir = Path.GetFullPath(args.Length == 2 ? args[1] :
    Path.Combine(Path.GetTempPath(), "excel-exportkit-benchmark-" + Guid.NewGuid().ToString("N")));
Directory.CreateDirectory(outputDir);
var input = Path.Combine(outputDir, "fixture.xlsx");

// Generate deterministic, realistic-ish mixed data with repeatable row count.
var generation = Stopwatch.StartNew();
using (var workbook = new XLWorkbook())
{
    var sheet = workbook.AddWorksheet("Benchmark");
    foreach (var (column, header) in new[]
    {
        (1, "Student"), (2, "Score"), (3, "Active"),
        (4, "ExamDate"), (5, "Notes"), (6, "Group")
    })
        sheet.Cell(1, column).Value = header;

    for (var i = 0; i < rows; i++)
    {
        var row = i + 2;
        sheet.Cell(row, 1).Value = $"Student {i + 1}";
        sheet.Cell(row, 2).Value = 800 + (i * 13) % 800;
        sheet.Cell(row, 3).Value = i % 3 == 0;
        sheet.Cell(row, 4).Value = new DateTime(2026, 1, 1).AddDays(i % 365);
        sheet.Cell(row, 5).Value = $"Term {i % 7}, practice \"{i % 5}\"";
        sheet.Cell(row, 6).Value = $"Group {i % 20}";
    }
    workbook.SaveAs(input);
}
generation.Stop();

var steps = new List<object>
{
    new
    {
        operation = "generate-xlsx",
        elapsedMs = generation.ElapsedMilliseconds,
        bytes = new FileInfo(input).Length,
        managedBytes = GC.GetTotalMemory(false)
    }
};

foreach (var format in rows > 10_000
             ? new[] { "csv", "json" }
             : new[] { "csv", "json", "pdf", "docx" })
{
    var path = Path.Combine(outputDir, "converted." + format);
    var options = new ExcelConversionOptions
    {
        WorksheetName = "Benchmark",
        PreserveJsonTypes = true,
        CsvProtectFormulas = true,
        PdfPreserveCellFormatting = format == "pdf",
        DocumentTitle = "Performance benchmark"
    };
    GC.Collect();
    GC.WaitForPendingFinalizers();
    var before = GC.GetAllocatedBytesForCurrentThread();
    var timer = Stopwatch.StartNew();
    switch (format)
    {
        case "csv": await ExcelConverter.ToCSVAsync(input, path, options); break;
        case "json": await ExcelConverter.ToJsonAsync(input, path, options); break;
        case "pdf": await ExcelConverter.ToPDFAsync(input, path, options); break;
        case "docx": await ExcelConverter.ToWordAsync(input, path, options); break;
    }
    timer.Stop();
    steps.Add(new
    {
        operation = "to-" + format,
        elapsedMs = timer.ElapsedMilliseconds,
        bytes = new FileInfo(path).Length,
        allocatedCurrentThreadBytes = Math.Max(0, GC.GetAllocatedBytesForCurrentThread() - before),
        managedBytes = GC.GetTotalMemory(false),
        peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64
    });
}

var results = new
{
    rows,
    columns = 6,
    framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    environment = Environment.OSVersion.ToString(),
    caveats = "Cold single-process benchmark; no reproducible cross-machine performance claims. Thread allocation excludes work on other threads.",
    steps
};
var json = JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true });
await File.WriteAllTextAsync(Path.Combine(outputDir, "benchmark.json"), json);
Console.WriteLine(json);
return 0;
