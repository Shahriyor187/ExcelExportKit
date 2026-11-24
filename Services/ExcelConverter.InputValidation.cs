using System.IO.Compression;

namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    /// <summary>
    /// Explicit opt-in safety limits. ClosedXML loads the entire workbook into
    /// memory; callers processing untrusted uploads should set these limits
    /// *before* loading it. Limits are disabled by default for compatibility.
    /// </summary>
    private static void ValidateWorkbookInput(string filePath, ExcelConversionOptions options, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(options);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        using var input = File.OpenRead(filePath);
        ValidateWorkbookInput(input, options, cancellationToken);
    }

    private static void ValidateWorkbookInput(Stream input, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(options);
        if (!input.CanRead || !input.CanSeek)
            throw new ArgumentException("Workbook input must be readable and seekable.", nameof(input));
        input.Position = 0;

        if (options.MaxWorkbookFileBytes is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxWorkbookFileBytes));
        if (options.MaxUncompressedWorkbookBytes is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxUncompressedWorkbookBytes));
        if (options.MaxWorkbookArchiveEntries is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxWorkbookArchiveEntries));
        if (options.MaxSelectedCells is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxSelectedCells));

        if (options.MaxWorkbookFileBytes is long maxFileBytes &&
            input.Length > maxFileBytes)
            throw new InvalidDataException(
                $"Workbook file exceeds the configured maximum of {maxFileBytes} bytes.");

        if (options.MaxUncompressedWorkbookBytes is null &&
            options.MaxWorkbookArchiveEntries is null)
            return;

        // A .xlsx is a ZIP archive. Inspect its entries before asking ClosedXML
        // to load the workbook. Metadata-only size checks are insufficient:
        // actually read entry streams to enforce expanded-byte limits even
        // when an archive has forged size metadata.
        try
        {
        using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
        if (options.MaxWorkbookArchiveEntries is int maxEntries &&
            archive.Entries.Count > maxEntries)
            throw new InvalidDataException(
                $"Workbook archive contains more than {maxEntries} entries.");

        if (options.MaxUncompressedWorkbookBytes is not long maxUncompressedBytes)
            return;

        var expanded = 0L;
        var buffer = new byte[64 * 1024];
        foreach (var entry in archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Length > maxUncompressedBytes - expanded)
                throw new InvalidDataException(
                    $"Workbook exceeds the configured expanded limit of {maxUncompressedBytes} bytes.");

            using var stream = entry.Open();
            expanded = CountExpandedBytes(stream, buffer, expanded, maxUncompressedBytes, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        }
        finally
        {
            input.Position = 0;
        }
    }

    // Shared by archive validation and deterministic stream-based tests.
    internal static long CountExpandedBytes(Stream stream, byte[] buffer, long expanded,
        long maxUncompressedBytes, CancellationToken cancellationToken)
    {
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var read = stream.Read(buffer, 0, buffer.Length);
            cancellationToken.ThrowIfCancellationRequested();
            if (read == 0)
                return expanded;
            if (read > maxUncompressedBytes - expanded)
                throw new InvalidDataException(
                    $"Workbook exceeds the configured expanded limit of {maxUncompressedBytes} bytes.");
            expanded += read;
        }
    }

    private static void ValidateSelectedCells(
        int firstRow, int lastRow, int firstCol, int lastCol, ExcelConversionOptions options)
    {
        if (options.MaxSelectedCells is not long maxCells)
            return;

        var rowCount = (long)lastRow - firstRow + 1;
        var colCount = (long)lastCol - firstCol + 1;
        if (rowCount <= 0 || colCount <= 0 ||
            rowCount * colCount > maxCells)
            throw new InvalidDataException(
                $"Selected worksheet exceeds the configured maximum of {maxCells} cells.");
    }
}
