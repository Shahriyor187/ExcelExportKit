using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using ExportFiles.Services;
using ClosedXML.Excel;

internal static class ProgressBenchmark
{
    private sealed class Counter : IProgress<ExportProgress>
    {
        public long Rows { get; private set; }
        public int Reports { get; private set; }
        public void Report(ExportProgress value)
        {
            Rows = value.RowsWritten;
            Reports++;
        }
    }

    private sealed record Trial(double Milliseconds, long AllocatedBytes, string Hash);

    internal static async Task RunAsync()
    {
        const int rows = 10000;
        const int interval = 1000;
        const int repeats = 5;
        using var fixture = new MemoryStream();
        using (var book = new XLWorkbook())
        {
            var sheet = book.AddWorksheet("Data");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Text";
            for (var i = 1; i <= rows; i++)
            {
                sheet.Cell(i + 1, 1).Value = i;
                sheet.Cell(i + 1, 2).Value = "Row " + i;
            }
            book.SaveAs(fixture);
        }
        var workbookBytes = fixture.ToArray();

        async Task<Trial> Measure(bool enabled)
        {
            var counter = enabled ? new Counter() : null;
            using var output = new MemoryStream();
            using var input = new MemoryStream(workbookBytes, writable: false);
            var options = new ExcelConversionOptions { Progress = counter, ProgressIntervalRows = interval };
            var allocated = GC.GetTotalAllocatedBytes(precise: true);
            var stopwatch = Stopwatch.StartNew();
            await ExcelConverter.ToCSVAsync(input, output, options);
            stopwatch.Stop();
            var bytes = GC.GetTotalAllocatedBytes(precise: true) - allocated;
            if (counter is not null && (counter.Rows != rows || counter.Reports != (rows + interval - 1) / interval))
                throw new InvalidOperationException("Unexpected progress report counts.");
            if (output.Length == 0) throw new InvalidOperationException("Empty benchmark output.");
            return new Trial(stopwatch.Elapsed.TotalMilliseconds, bytes,
                Convert.ToHexString(SHA256.HashData(output.GetBuffer().AsSpan(0, checked((int)output.Length)))));
        }

        // Warm up both paths; alternate ordering to reduce systematic bias.
        var reference = await Measure(false);
        var warmup = await Measure(true);
        if (reference.Hash != warmup.Hash) throw new InvalidOperationException("Progress changed output.");
        var off = new List<Trial>();
        var on = new List<Trial>();
        for (var i = 0; i < repeats; i++)
        {
            if (i % 2 == 0) { off.Add(await Measure(false)); on.Add(await Measure(true)); }
            else { on.Add(await Measure(true)); off.Add(await Measure(false)); }
        }
        if (off.Concat(on).Any(t => t.Hash != reference.Hash))
            throw new InvalidOperationException("Benchmark output differs between trials.");
        double Median(IEnumerable<double> values) => values.Order().ElementAt(repeats / 2);
        var offMs = Median(off.Select(t => t.Milliseconds));
        var onMs = Median(on.Select(t => t.Milliseconds));
        var offAlloc = Median(off.Select(t => (double)t.AllocatedBytes));
        var onAlloc = Median(on.Select(t => (double)t.AllocatedBytes));
        Console.WriteLine("PROGRESS_BENCHMARK " + JsonSerializer.Serialize(new
        {
            rows, interval, repeats,
            disabledMedianMs = offMs, enabledMedianMs = onMs,
            timeOverheadPercent = (onMs / offMs - 1) * 100,
            disabledMedianAllocatedBytes = offAlloc, enabledMedianAllocatedBytes = onAlloc,
            allocationOverheadPercent = (onAlloc / offAlloc - 1) * 100,
            disabledTrialsMs = off.Select(t => t.Milliseconds),
            enabledTrialsMs = on.Select(t => t.Milliseconds),
            outputIdentical = true
        }));
        // Shared CI runners are noisy; report measurements rather than enforce
        // a wall-clock ratio. Exact counts and byte equivalence are hard gates.
    }
}
