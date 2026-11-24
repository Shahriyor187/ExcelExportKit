namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    // A lease owns only buffers created by the library. Caller streams stay open.
    private sealed class WorkbookInputLease(Stream stream, bool ownsStream, string? temporaryPath = null) : IAsyncDisposable
    {
        public Stream Stream { get; } = stream;

        public async ValueTask DisposeAsync()
        {
            try
            {
                if (ownsStream)
                    await Stream.DisposeAsync().ConfigureAwait(false);
            }
            finally
            {
                if (temporaryPath is not null)
                    File.Delete(temporaryPath);
            }
        }
    }

    private static async Task<WorkbookInputLease> PrepareWorkbookInputAsync(
        Stream input, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(options);
        if (!input.CanRead)
            throw new ArgumentException("Workbook input must be readable.", nameof(input));
        if (!Enum.IsDefined(options.InputBufferingMode))
            throw new ArgumentOutOfRangeException(nameof(options.InputBufferingMode));
        if (options.MaxBufferedInputBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxBufferedInputBytes));
        if (options.MaxWorkbookFileBytes is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxWorkbookFileBytes));

        if (input.CanSeek)
        {
            if (input.Position != 0)
                throw new ArgumentException("Seekable workbook input must be positioned at zero.", nameof(input));
            if (options.MaxWorkbookFileBytes is long max && input.Length > max)
                throw new InvalidDataException($"Workbook exceeds the configured maximum of {max} bytes.");
            return new WorkbookInputLease(input, false);
        }

        if (options.InputBufferingMode == InputBufferingMode.RejectNonSeekable)
            throw new NotSupportedException("Non-seekable workbook input requires an explicit buffering mode.");

        var limit = Math.Min(options.MaxBufferedInputBytes, options.MaxWorkbookFileBytes ?? long.MaxValue);
        WorkbookInputLease lease;
        if (options.InputBufferingMode == InputBufferingMode.Memory)
        {
            lease = new WorkbookInputLease(new MemoryStream(), true);
        }
        else
        {
            var directory = options.TemporaryDirectory ?? Path.GetTempPath();
            var path = Path.Combine(directory, $".excel-input-{Guid.NewGuid():N}.xlsx");
            var fileOptions = new FileStreamOptions
            {
                Mode = FileMode.CreateNew,
                Access = FileAccess.ReadWrite,
                Share = FileShare.None,
                Options = FileOptions.Asynchronous | FileOptions.DeleteOnClose,
                BufferSize = 64 * 1024
            };
            if (!OperatingSystem.IsWindows())
                fileOptions.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            var file = new FileStream(path, fileOptions);
            lease = new WorkbookInputLease(file, true);
        }

        try
        {
            await CopyWorkbookInputAsync(input, lease.Stream, limit, cancellationToken).ConfigureAwait(false);
            lease.Stream.Position = 0;
            return lease;
        }
        catch
        {
            await lease.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    // Read at most one byte beyond the limit to distinguish exact-sized input
    // from oversized input. Never write bytes that exceed the configured limit.
    internal static async Task CopyWorkbookInputAsync(Stream input, Stream output,
        long limit, CancellationToken cancellationToken)
    {
        if (limit <= 0) throw new ArgumentOutOfRangeException(nameof(limit));
        var buffer = new byte[64 * 1024];
        long copied = 0;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var remaining = limit - copied;
            var request = remaining >= buffer.Length ? buffer.Length : (int)remaining + 1;
            var read = await input.ReadAsync(buffer.AsMemory(0, request), cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (read == 0) break;
            if (read > remaining)
                throw new InvalidDataException($"Workbook input exceeds the configured buffering limit of {limit} bytes.");
            await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            copied += read;
        }
        await output.FlushAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }
}
