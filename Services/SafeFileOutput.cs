using System.Diagnostics;

namespace ExportFiles.Services;

/// <summary>Details of a successfully published file. Elapsed includes rendering and publication.</summary>
/// <param name="OutputPath">Absolute destination path.</param>
/// <param name="OutputBytes">Final file length in bytes.</param>
/// <param name="Elapsed">Elapsed time measured for the operation.</param>
public sealed record ConversionFileResult(string OutputPath, long OutputBytes, TimeSpan Elapsed);

internal static class SafeFileOutput
{
    internal static async Task<ConversionFileResult> WriteAsync(
        string outputFile, long? maxOutputBytes, Func<Stream, CancellationToken, Task> write,
        CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFile);
        if (maxOutputBytes is <= 0)
            throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));
        cancellationToken.ThrowIfCancellationRequested();
        var clock = Stopwatch.StartNew();
        var fullPath = Path.GetFullPath(outputFile);
        var parent = Path.GetDirectoryName(fullPath)!;
        if (!Directory.Exists(parent))
            throw new DirectoryNotFoundException("The destination directory must exist.");

        var directory = Path.Combine(parent, ".excel-safe-" + Guid.NewGuid().ToString("N"));
        if (OperatingSystem.IsWindows())
            Directory.CreateDirectory(directory);
        else
            Directory.CreateDirectory(directory, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        var stage = Path.Combine(directory, "output");
        try
        {
            long bytes;
            await using (var output = new FileStream(stage, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.None, 8192, useAsync: true))
            {
                await write(output, cancellationToken).ConfigureAwait(false);
                await output.FlushAsync(cancellationToken).ConfigureAwait(false);
                bytes = output.Length;
            }
            if (maxOutputBytes is long limit && bytes > limit)
                throw new IOException($"Final output size {bytes} bytes exceeds the limit of {limit} bytes.");

            cancellationToken.ThrowIfCancellationRequested();
            // Same-filesystem rename is the publication boundary. Do not check cancellation after it.
            File.Move(stage, fullPath, overwrite: true);
            return new ConversionFileResult(fullPath, bytes, clock.Elapsed);
        }
        finally
        {
            // Best-effort cleanup must not mask a renderer failure or report a published file as failed.
            try { Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
