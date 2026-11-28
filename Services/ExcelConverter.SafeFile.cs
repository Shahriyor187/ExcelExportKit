namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    /// <summary>
    /// Converts a workbook to a staged file using the stream API's input safety defaults.
    /// Supported formats are csv, tsv, json, pdf and docx (case insensitive).
    /// The optional positive size limit is a final-file check, not a temporary storage or memory quota.
    /// The destination remains unchanged on failure or cancellation before publication.
    /// </summary>
    public static Task<ConversionFileResult> ConvertFileAsync(string inputFile, string outputFile,
        string format, ExcelConversionOptions? options = null, long? maxOutputBytes = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputFile);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFile);
        ArgumentException.ThrowIfNullOrWhiteSpace(format);
        if (string.Equals(Path.GetFullPath(inputFile), Path.GetFullPath(outputFile),
            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            throw new ArgumentException("Input and output paths must be distinct.", nameof(outputFile));

        Func<Stream, CancellationToken, Task> write = format.ToLowerInvariant() switch
        {
            "csv" => (output, token) => ToCSVAsync(inputFile, output, options, token),
            "tsv" => (output, token) => ToTSVAsync(inputFile, output, options, token),
            "json" => (output, token) => ToJsonAsync(inputFile, output, options, token),
            "pdf" => (output, token) => ToPDFAsync(inputFile, output, options, token),
            "docx" => (output, token) => ToWordAsync(inputFile, output, options, token),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };
        return SafeFileOutput.WriteAsync(outputFile, maxOutputBytes, write, cancellationToken);
    }
}
