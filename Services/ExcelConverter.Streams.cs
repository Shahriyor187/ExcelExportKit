namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    private enum StreamFormat { Csv, Tsv, Json, Pdf, Word }

    /// <summary>Converts a workbook stream to CSV, leaving caller streams open.</summary>
    public static Task ToCSVAsync(Stream input, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamAsync(input, output, StreamFormat.Csv, options, cancellationToken);

    /// <summary>Converts a workbook file to a caller-owned CSV output stream.</summary>
    public static Task ToCSVAsync(string inputFile, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertPathToStreamAsync(inputFile, output, StreamFormat.Csv, options, cancellationToken);

    /// <summary>Converts a workbook stream to a CSV file with staged publication.</summary>
    public static Task ToCSVAsync(Stream input, string outputFile, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamToPathAsync(input, outputFile, StreamFormat.Csv, options, cancellationToken);

    /// <summary>Converts a workbook stream to TSV, leaving caller streams open.</summary>
    public static Task ToTSVAsync(Stream input, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamAsync(input, output, StreamFormat.Tsv, options, cancellationToken);

    /// <summary>Converts a workbook file to a caller-owned TSV output stream.</summary>
    public static Task ToTSVAsync(string inputFile, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertPathToStreamAsync(inputFile, output, StreamFormat.Tsv, options, cancellationToken);

    /// <summary>Converts a workbook stream to a TSV file with staged publication.</summary>
    public static Task ToTSVAsync(Stream input, string outputFile, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamToPathAsync(input, outputFile, StreamFormat.Tsv, options, cancellationToken);

    /// <summary>Converts a workbook stream to Json, leaving caller streams open.</summary>
    public static Task ToJsonAsync(Stream input, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamAsync(input, output, StreamFormat.Json, options, cancellationToken);

    /// <summary>Converts a workbook file to a caller-owned Json output stream.</summary>
    public static Task ToJsonAsync(string inputFile, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertPathToStreamAsync(inputFile, output, StreamFormat.Json, options, cancellationToken);

    /// <summary>Converts a workbook stream to a Json file with staged publication.</summary>
    public static Task ToJsonAsync(Stream input, string outputFile, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamToPathAsync(input, outputFile, StreamFormat.Json, options, cancellationToken);

    /// <summary>Converts a workbook stream to PDF, leaving caller streams open.</summary>
    public static Task ToPDFAsync(Stream input, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamAsync(input, output, StreamFormat.Pdf, options, cancellationToken);

    /// <summary>Converts a workbook file to a caller-owned PDF output stream.</summary>
    public static Task ToPDFAsync(string inputFile, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertPathToStreamAsync(inputFile, output, StreamFormat.Pdf, options, cancellationToken);

    /// <summary>Converts a workbook stream to a PDF file with staged publication.</summary>
    public static Task ToPDFAsync(Stream input, string outputFile, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamToPathAsync(input, outputFile, StreamFormat.Pdf, options, cancellationToken);

    /// <summary>Converts a workbook stream to Word, leaving caller streams open.</summary>
    public static Task ToWordAsync(Stream input, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamAsync(input, output, StreamFormat.Word, options, cancellationToken);

    /// <summary>Converts a workbook file to a caller-owned Word output stream.</summary>
    public static Task ToWordAsync(string inputFile, Stream output, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertPathToStreamAsync(inputFile, output, StreamFormat.Word, options, cancellationToken);

    /// <summary>Converts a workbook stream to a Word file with staged publication.</summary>
    public static Task ToWordAsync(Stream input, string outputFile, ExcelConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        ConvertStreamToPathAsync(input, outputFile, StreamFormat.Word, options, cancellationToken);

    private static string CreatePrivateStageDirectory(string parent)
    {
        // The caller must supply a trusted directory, not a path controlled by uploads.
        var path = Path.Combine(parent, ".excel-stage-" + Guid.NewGuid().ToString("N"));
        if (OperatingSystem.IsWindows())
            Directory.CreateDirectory(path);
        else
            Directory.CreateDirectory(path, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        return path;
    }

    private static async Task ConvertPathToStreamAsync(string inputFile, Stream output, StreamFormat format,
        ExcelConversionOptions? options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentException.ThrowIfNullOrWhiteSpace(inputFile);
        await using var input = File.OpenRead(inputFile);
        await ConvertStreamAsync(input, output, format, options, cancellationToken).ConfigureAwait(false);
    }

    private static async Task ConvertStreamToPathAsync(Stream input, string outputFile, StreamFormat format,
        ExcelConversionOptions? options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(input);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFile);
        var fullPath = Path.GetFullPath(outputFile);
        var stageDirectory = CreatePrivateStageDirectory(Path.GetDirectoryName(fullPath)!);
        try
        {
            var stage = Path.Combine(stageDirectory, "output");
            await using (var output = new FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None, 8192, useAsync: true))
                await ConvertStreamAsync(input, output, format, options, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(stage, fullPath, overwrite: true);
        }
        finally
        {
            Directory.Delete(stageDirectory, recursive: true);
        }
    }

    private static async Task ConvertStreamAsync(Stream input, Stream output, StreamFormat format,
        ExcelConversionOptions? options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(output);
        if (ReferenceEquals(input, output))
            throw new ArgumentException("Input and output streams must be distinct.", nameof(output));
        if (!output.CanWrite)
            throw new ArgumentException("Output stream must be writable.", nameof(output));
        if (output.CanSeek && (output.Position != 0 || output.Length != 0))
            throw new ArgumentException("Seekable output must be empty and positioned at zero.", nameof(output));

        options = (options ?? new ExcelConversionOptions()).WithStreamSafetyDefaults();
        ValidateRowWindow(options);
        ValidateInputLimits(options);
        if (format == StreamFormat.Csv) ValidateCsvDelimiter(options.CsvDelimiter);
        if (format == StreamFormat.Pdf && options.PdfIncludeEmbeddedImages && !options.PdfPreserveCellFormatting)
            throw new ArgumentException("PdfIncludeEmbeddedImages requires PdfPreserveCellFormatting.", nameof(options));

        await using var lease = await PrepareWorkbookInputAsync(input, options, cancellationToken).ConfigureAwait(false);
        var workbook = lease.Stream;
        switch (format)
        {
            case StreamFormat.Csv:
            case StreamFormat.Tsv:
                await WriteSelectedDelimitedAsync(ReadWorksheet(workbook, options, cancellationToken), output, options,
                    format == StreamFormat.Tsv ? '\t' : options.CsvDelimiter, cancellationToken).ConfigureAwait(false);
                break;
            case StreamFormat.Json:
                await WriteSelectedJsonAsync(ReadWorksheet(workbook, options, cancellationToken), output, options, cancellationToken).ConfigureAwait(false);
                break;
            case StreamFormat.Pdf:
            case StreamFormat.Word:
                // The document renderers require a file destination here. Keep
                // intermediate output private, and copy only after successful rendering.
                var directory = CreatePrivateStageDirectory(options.TemporaryDirectory ?? Path.GetTempPath());
                try
                {
                    var path = Path.Combine(directory, format == StreamFormat.Pdf ? "output.pdf" : "output.docx");
                    if (format == StreamFormat.Word)
                    {
                        var sheet = ReadWorksheet(workbook, options, cancellationToken);
                        await Task.Run(() => WriteStyledWordDocument(sheet, path, options, cancellationToken), cancellationToken).ConfigureAwait(false);
                    }
                    else if (options.PdfUseWorksheetPrintSettings)
                        await RenderWorksheetPrintPdfAsync(ReadWorksheetPrintLayout(workbook, options, cancellationToken), path, options, cancellationToken).ConfigureAwait(false);
                    else if (options.PdfPreserveCellFormatting)
                        await RenderFormattingAwarePdfAsync(ReadPdfFormatting(workbook, options, cancellationToken), path, options, cancellationToken).ConfigureAwait(false);
                    else
                        await RenderSelectedPdfAsync(ReadWorksheet(workbook, options, cancellationToken), path, options, cancellationToken).ConfigureAwait(false);

                    cancellationToken.ThrowIfCancellationRequested();
                    await using var rendered = File.OpenRead(path);
                    await rendered.CopyToAsync(output, 64 * 1024, cancellationToken).ConfigureAwait(false);
                }
                finally
                {
                    Directory.Delete(directory, recursive: true);
                }
                break;
        }
        cancellationToken.ThrowIfCancellationRequested();
        await output.FlushAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }
}
