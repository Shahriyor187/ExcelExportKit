using System.IO.Compression;

namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    /// <summary>
    /// Explicit opt-in safety limits. ClosedXML loads the entire workbook into
    /// memory; callers processing untrusted uploads should set these limits
    /// *before* loading it. Limits are disabled by default for compatibility.
    /// </summary>
    private static void ValidateWorkbookInput(string filePath, ExcelConversionOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        if (options.MaxWorkbookFileBytes is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxWorkbookFileBytes));
        if (options.MaxUncompressedWorkbookBytes is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxUncompressedWorkbookBytes));
        if (options.MaxWorkbookArchiveEntries is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxWorkbookArchiveEntries));
        if (options.MaxSelectedCells is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxSelectedCells));

        var file = new FileInfo(filePath);
        if (!file.Exists)
            throw new FileNotFoundException("The Excel input file was not found.", filePath);

        if (options.MaxWorkbookFileBytes is long maxFileBytes &&
            file.Length > maxFileBytes)
            throw new InvalidDataException(
                $"Workbook file exceeds the configured maximum of {maxFileBytes} bytes.");

        if (options.MaxUncompressedWorkbookBytes is null &&
            options.MaxWorkbookArchiveEntries is null)
            return;

        // A .xlsx is a ZIP archive. Inspect its entries before asking ClosedXML
        // to load the workbook. Metadata-only size checks are insufficient:
        // actually read entry streams to enforce expanded-byte limits even
        // when an archive has forged size metadata.
        using var archive = ZipFile.OpenRead(filePath);
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
            if (entry.Length > maxUncompressedBytes - expanded)
                throw new InvalidDataException(
                    $"Workbook exceeds the configured expanded limit of {maxUncompressedBytes} bytes.");

            using var stream = entry.Open();
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) != 0)
            {
                if (read > maxUncompressedBytes - expanded)
                    throw new InvalidDataException(
                        $"Workbook exceeds the configured expanded limit of {maxUncompressedBytes} bytes.");
                expanded += read;
            }
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
