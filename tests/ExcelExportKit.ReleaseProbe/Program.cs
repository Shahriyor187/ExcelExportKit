using System.Diagnostics;
using System.Text.Json;
using ExportFiles.Services;

// Local-only diagnostic for release validation on a representative workbook.
// Never prints worksheet cells, workbook path, sheet name or exception messages.
// No HTTP/API calls and no source workbook copied into the result report.

if (args.Length < 1 || args[0] is "--help" or "-h")
{
    Console.WriteLine("Usage: dotnet run -c Release --project tests/ExcelExportKit.ReleaseProbe -- <workbook.xlsx> " +
        "[--sheet NAME] [--range A1:F100] [--formats csv,json,pdf,docx] [--report result.json]");
    Console.WriteLine("Default formats: csv,json. Only anonymized diagnostic metadata is reported; " +
        "conversions are deleted from the local temp directory.");
    return args.Length == 1 ? 0 : 2;
}

var workbookPath = args[0];
string? sheetName = null, rangeAddress = null, reportPath = null;
var formats = "csv,json";
for (var i = 1; i < args.Length; i++)
{
    if (i + 1 >= args.Length)
    {
        Console.Error.WriteLine("Missing value for option.");
        return 2;
    }
    var value = args[++i];
    switch (args[i - 1])
    {
        case "--sheet": sheetName = value; break;
        case "--range": rangeAddress = value; break;
        case "--formats": formats = value; break;
        case "--report": reportPath = value; break;
        default:
            Console.Error.WriteLine("Unknown option. See --help.");
            return 2;
    }
}

var selected = formats.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
    .Select(x => x.ToLowerInvariant()).Distinct().ToArray();
if (selected.Length is < 1 or > 4 ||
    selected.Any(x => x is not ("csv" or "json" or "pdf" or "docx")))
{
    Console.Error.WriteLine("Select one or more of csv,json,pdf,docx.");
    return 2;
}
if (string.IsNullOrWhiteSpace(workbookPath) ||
    !string.Equals(Path.GetExtension(workbookPath), ".xlsx", StringComparison.OrdinalIgnoreCase) ||
    !File.Exists(workbookPath))
{
    Console.Error.WriteLine("Input must be an existing .xlsx file.");
    return 2;
}

var folder = Path.Combine(Path.GetTempPath(), "excel-exportkit-probe-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(folder);
try
{
    var options = new ExcelConversionOptions
    {
        WorksheetName = sheetName,
        RangeAddress = rangeAddress,
        SkipEmptyRows = true,
        PreserveJsonTypes = true,
        PdfPreserveCellFormatting = true,
        DocumentTitle = "Local conversion probe",
        MaxWorkbookFileBytes = 20L * 1024 * 1024,
        MaxUncompressedWorkbookBytes = 100L * 1024 * 1024,
        MaxWorkbookArchiveEntries = 2048,
        MaxSelectedCells = 250_000
    };
    var measurements = new List<object>();
    var failures = 0;
    foreach (var format in selected)
    {
        var output = Path.Combine(folder, "converted." + format);
        var clock = Stopwatch.StartNew();
        string? errorCategory = null;
        long outputBytes = 0;
        try
        {
            switch (format)
            {
                case "csv": await ExcelConverter.ToCSVAsync(workbookPath, output, options); break;
                case "json": await ExcelConverter.ToJsonAsync(workbookPath, output, options); break;
                case "pdf": await ExcelConverter.ToPDFAsync(workbookPath, output, options); break;
                case "docx": await ExcelConverter.ToWordAsync(workbookPath, output, options); break;
            }
            var file = new FileInfo(output);
            if (!file.Exists || file.Length == 0)
                throw new InvalidDataException("Missing or empty output.");
            outputBytes = file.Length;
            if (format == "json")
            {
                await using var jsonFile = File.OpenRead(output);
                using var jsonDocument = await JsonDocument.ParseAsync(jsonFile);
                if (jsonDocument.RootElement.ValueKind != JsonValueKind.Array)
                    throw new InvalidDataException("Invalid JSON array output.");
            }
            if (format == "pdf")
            {
                using var stream = File.OpenRead(output);
                Span<byte> signature = stackalloc byte[5];
                if (stream.Read(signature) != 5 || System.Text.Encoding.ASCII.GetString(signature) != "%PDF-")
                    throw new InvalidDataException("Invalid PDF output.");
            }
            if (format == "docx")
            {
                using var archive = System.IO.Compression.ZipFile.OpenRead(output);
                if (!archive.Entries.Any(entry => entry.FullName == "word/document.xml"))
                    throw new InvalidDataException("Invalid DOCX document package.");
            }
        }
        catch (Exception ex)
        {
            // Keep error messages and input values private; only type is included.
            errorCategory = ex.GetType().Name;
            failures++;
        }
        clock.Stop();
        Process.GetCurrentProcess().Refresh();
        measurements.Add(new
        {
            format,
            success = errorCategory is null,
            elapsedMilliseconds = Math.Round(clock.Elapsed.TotalMilliseconds, 1),
            outputBytes,
            errorCategory
        });
        if (File.Exists(output)) File.Delete(output);
    }

    var report = new
    {
        schemaVersion = 1,
        note = "Local-only probe. No workbook content, filename, sheet name or exception message is recorded.",
        framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
        peakProcessWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64,
        peakMemoryCaveat = "Cumulative process peak across selected formats; not per-format isolation.",
        formats = measurements,
        overallSuccess = failures == 0
    };
    var json = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
    if (reportPath is not null)
    {
        var destination = Path.GetFullPath(reportPath);
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        await File.WriteAllTextAsync(destination, json);
    }
    Console.WriteLine(json);
    return failures == 0 ? 0 : 1;
}
finally
{
    // Never retain converted workbook data after the probe.
    Directory.Delete(folder, recursive: true);
}
