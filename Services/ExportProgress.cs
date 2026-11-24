using System.Diagnostics;

namespace ExportFiles.Services;

/// <summary>Data rows processed by an export, excluding headers and skipped records.
/// This is not a guarantee that the destination has been finalized.</summary>
public readonly struct ExportProgress
{
    /// <summary>Number of data rows processed so far, excluding headers.</summary>
    public long RowsWritten { get; init; }
    /// <summary>Known or estimated output row count; null if unknown.</summary>
    public long? TotalRows { get; init; }
    /// <summary>Monotonic elapsed time since row processing began.</summary>
    public TimeSpan Elapsed { get; init; }
}

internal sealed class ExportProgressTracker
{
    private readonly IProgress<ExportProgress> progress;
    private readonly int interval;
    private readonly long? total;
    private readonly CancellationToken token;
    private readonly Stopwatch clock = Stopwatch.StartNew();
    private long rows;
    private long lastReported = -1;

    private ExportProgressTracker(IProgress<ExportProgress> progress, int interval, long? total, CancellationToken token)
    {
        this.progress = progress;
        this.interval = interval;
        this.total = total;
        this.token = token;
    }

    internal static void Validate(int interval, long? total)
    {
        if (interval <= 0) throw new ArgumentOutOfRangeException("ProgressIntervalRows");
        if (total is < 0) throw new ArgumentOutOfRangeException("TotalRows");
    }

    internal static ExportProgressTracker? Create(IProgress<ExportProgress>? progress, int interval,
        long? total, CancellationToken token = default)
    {
        Validate(interval, total);
        return progress is null ? null : new ExportProgressTracker(progress, interval, total, token);
    }

    internal void RowWritten()
    {
        token.ThrowIfCancellationRequested();
        rows++;
        if (rows % interval == 0) Report();
    }

    internal void Finish()
    {
        token.ThrowIfCancellationRequested();
        if (lastReported != rows) Report();
    }

    private void Report()
    {
        token.ThrowIfCancellationRequested();
        // Caller estimates can be wrong. Never report a total smaller than
        // the number of rows already processed.
        progress.Report(new ExportProgress
        {
            RowsWritten = rows,
            TotalRows = total.HasValue && total.Value >= rows ? total : null,
            Elapsed = clock.Elapsed
        });
        lastReported = rows;
        token.ThrowIfCancellationRequested();
    }
}
