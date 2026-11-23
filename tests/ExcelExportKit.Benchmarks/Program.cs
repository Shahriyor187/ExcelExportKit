using System.Diagnostics;
using System.Text.Json;
using ClosedXML.Excel;
using ExportFiles.Services;

// Worker modes are called by run_benchmarks.py. Each conversion runs in a
// separate OS process and is warmed up before measurement.
if (args.Length == 3 && args[0] == "generate" &&
    int.TryParse(args[1], out var rowCount) && rowCount is >= 1 and <= 100_000)
{
    var input = Path.GetFullPath(args[2]);
    Directory.CreateDirectory(Path.GetDirectoryName(input)!);
    using (var workbook = new XLWorkbook())
    {
        var sheet = workbook.AddWorksheet("Benchmark");
        string[] headers = ["Student", "Score", "Active", "ExamDate", "Notes", "Group"];
        for (var i = 0; i < headers.Length; i++)
            sheet.Cell(1, i + 1).Value = headers[i];

        for (var i = 0; i < rowCount; i++)
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

    Console.WriteLine(JsonSerializer.Serialize(new
    {
        operation = "generate",
        rows = rowCount,
        bytes = new FileInfo(input).Length
    }));
    return 0;
}

if (args.Length == 4 && args[0] == "measure" &&
    args[1] is "csv" or "json" or "pdf" or "docx")
{
    var format = args[1];
    var input = Path.GetFullPath(args[2]);
    var output = Path.GetFullPath(args[3]);
    Directory.CreateDirectory(Path.GetDirectoryName(output)!);
    var options = new ExcelConversionOptions
    {
        WorksheetName = "Benchmark",
        PreserveJsonTypes = format == "json",
        CsvProtectFormulas = true,
        PdfPreserveCellFormatting = format == "pdf",
        DocumentTitle = "Performance benchmark"
    };

    async Task Convert(string path)
    {
        switch (format)
        {
            case "csv": await ExcelConverter.ToCSVAsync(input, path, options); break;
            case "json": await ExcelConverter.ToJsonAsync(input, path, options); break;
            case "pdf": await ExcelConverter.ToPDFAsync(input, path, options); break;
            case "docx": await ExcelConverter.ToWordAsync(input, path, options); break;
        }
    }

    // Warm up the same format in-process, then take one measured conversion.
    // The Python runner repeats each isolated worker three times.
    var warmup = output + ".warmup";
    try
    {
        await Convert(warmup);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var beforeAllocation = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        await Convert(output);
        stopwatch.Stop();
        var allocation = GC.GetTotalAllocatedBytes(precise: true) - beforeAllocation;
        var process = Process.GetCurrentProcess();
        process.Refresh();

        Console.WriteLine(JsonSerializer.Serialize(new
        {
            operation = format,
            elapsedMs = stopwatch.Elapsed.TotalMilliseconds,
            outputBytes = new FileInfo(output).Length,
            allocatedBytes = allocation,
            managedHeapBytes = GC.GetTotalMemory(false),
            workingSetBytes = process.WorkingSet64,
            peakWorkingSetBytes = process.PeakWorkingSet64
        }));
    }
    finally
    {
        if (File.Exists(warmup))
            File.Delete(warmup);
    }
    return 0;
}

Console.Error.WriteLine(
    "Usage: dotnet ExcelExportKit.Benchmarks.dll generate <rows> <fixture.xlsx>\n" +
    "   or: dotnet ExcelExportKit.Benchmarks.dll measure <csv|json|pdf|docx> <fixture.xlsx> <output>");
return 2;
