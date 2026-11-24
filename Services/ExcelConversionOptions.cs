namespace ExportFiles.Services;

/// <summary>Explicit policy for non-seekable workbook input streams.</summary>
public enum InputBufferingMode
{
    /// <summary>Reject non-seekable input without reading it.</summary>
    RejectNonSeekable,
    /// <summary>Buffer non-seekable input in memory, subject to the configured byte limit.</summary>
    Memory,
    /// <summary>Buffer non-seekable input in a private temporary file, subject to the configured byte limit.</summary>
    TemporaryFile
}

/// <summary>Settings for worksheet selection and conversion output.</summary>
public sealed class ExcelConversionOptions
{
    /// <summary>Optional row progress notifications. Await the export itself to determine success.</summary>
    public IProgress<ExportProgress>? Progress { get; init; }
    /// <summary>Number of data rows between progress notifications.</summary>
    public int ProgressIntervalRows { get; init; } = 1000;

    /// <summary>Policy for non-seekable workbook input. Seekable input needs no buffering.</summary>
    public InputBufferingMode InputBufferingMode { get; init; } = InputBufferingMode.RejectNonSeekable;

    /// <summary>Maximum bytes buffered from non-seekable input. Default is 64 MiB.</summary>
    public long MaxBufferedInputBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>Optional existing directory for library-owned temporary files. Null uses the system temporary directory.</summary>
    public string? TemporaryDirectory { get; init; }

    // New stream APIs use finite input limits by default without altering legacy path APIs.
    internal ExcelConversionOptions WithStreamSafetyDefaults()
    {
        var copy = (ExcelConversionOptions)MemberwiseClone();
        copy.MaxWorkbookFileBytes ??= 64L * 1024 * 1024;
        copy.MaxUncompressedWorkbookBytes ??= 256L * 1024 * 1024;
        copy.MaxWorkbookArchiveEntries ??= 4096;
        copy.MaxSelectedCells ??= 1_000_000;
        return copy;
    }


    /// <summary>Number of physical data rows to skip after the header, before empty-row filtering.</summary>
    public long RowOffset { get; init; }
    /// <summary>Maximum physical data rows to select, or null for unlimited. Headers are excluded.</summary>
    public long? MaxRows { get; init; }

    /// <summary>Select a worksheet by name. Takes precedence over WorksheetIndex.</summary>
    public string? WorksheetName { get; set; }

    /// <summary>Zero-based worksheet index; defaults to the first worksheet.</summary>
    public int WorksheetIndex { get; set; }

    /// <summary>
    /// Optional A1-style range (for example "B3:D100"). When supplied, the first row
    /// of this range is treated as headers and only its cells are converted.
    /// </summary>
    public string? RangeAddress { get; set; }

    /// <summary>Skip rows whose selected cells contain no values.</summary>
    public bool SkipEmptyRows { get; set; }

    /// <summary>Emit JSON numbers and booleans using their native types when possible.</summary>
    public bool PreserveJsonTypes { get; set; }

    /// <summary>CSV field separator.</summary>
    public char CsvDelimiter { get; set; } = ',';

    /// <summary>Quote strings to prevent spreadsheet formula injection when reopening CSV exports.</summary>
    public bool CsvProtectFormulas { get; set; } = true;

    /// <summary>Title written above exported PDF and Word tables.</summary>
    public string DocumentTitle { get; set; } = "Excel worksheet export";

    /// <summary>Use landscape A4 orientation for PDF documents.</summary>
    public bool PdfLandscape { get; set; }

    /// <summary>
    /// Opt in to worksheet print settings for formatting-aware PDF output:
    /// rectangular print areas, margins, landscape orientation, repeated title rows,
    /// scaling percentage and manual horizontal/vertical page breaks.
    /// Incompatible cropped merges are rejected. Excel's exact print engine,
    /// fit-to-N-pages and variable paper sizes are not emulated.
    /// Requires PdfPreserveCellFormatting; disabled by default.
    /// </summary>
    public bool PdfUseWorksheetPrintSettings { get; set; }

    /// <summary>
    /// Optional 10-400% override of the worksheet's print scaling percentage.
    /// Applies only when PdfUseWorksheetPrintSettings is enabled. Scaling changes
    /// PDF cell text, padding, borders and image sizes, not Excel's fit-to-page logic.
    /// </summary>
    public int? PdfPrintScalePercent { get; set; }

    // Internal-only: a print job may split a sheet into sections. An image
    // anchored outside the current section belongs to another page, not an error.
    internal bool PdfIgnoreImagesOutsideSelectedRange { get; set; }

    /// <summary>
    /// Optional maximum input XLSX file size in bytes, checked before loading it.
    /// Set for untrusted uploads; null leaves existing behavior unchanged.
    /// </summary>
    public long? MaxWorkbookFileBytes { get; set; }

    /// <summary>
    /// Optional maximum total expanded ZIP entry bytes for the XLSX archive.
    /// The archive is read once before ClosedXML loads it to enforce this limit.
    /// Null disables the check for backward compatibility.
    /// </summary>
    public long? MaxUncompressedWorkbookBytes { get; set; }

    /// <summary>Optional maximum number of XLSX ZIP archive entries.</summary>
    public int? MaxWorkbookArchiveEntries { get; set; }

    /// <summary>
    /// Optional maximum number of cells in the selected worksheet rectangle.
    /// Prevents accidentally processing enormous sparse ranges.
    /// </summary>
    public long? MaxSelectedCells { get; set; }

    /// <summary>Word table style identifier.</summary>
    public string WordTableStyle { get; set; } = "TableGrid";

    /// <summary>
    /// Preserve supported Excel fill/font colors, bold text, horizontal alignment and column-width proportions
    /// when exporting to PDF. Off by default to preserve the existing PDF appearance.
    /// Fully selected merged cells, solid approximations of borders, row heights and font sizes are supported. Embedded images are separately opt-in; exact Excel print-layout fidelity is not supported.
    /// </summary>
    public bool PdfPreserveCellFormatting { get; set; }

    /// <summary>
    /// Embed PNG/JPEG worksheet pictures anchored to the selected cells when
    /// PdfPreserveCellFormatting is enabled. Pictures are fitted inside their
    /// anchor cell rather than positioned as floating Excel objects.
    /// Disabled by default for backward compatibility.
    /// </summary>
    public bool PdfIncludeEmbeddedImages { get; set; }

    /// <summary>
    /// Maximum permitted compressed bytes for one embedded image (4 MiB by default).
    /// This limit prevents loading arbitrarily large pictures into the renderer.
    /// </summary>
    public int PdfMaxEmbeddedImageBytes { get; set; } = 4 * 1024 * 1024;

    /// <summary>
    /// Maximum number of cell-anchored pictures to embed (32 by default).
    /// </summary>
    public int PdfMaxEmbeddedImages { get; set; } = 32;

    /// <summary>Use a landscape page in Word exports.</summary>
    public bool WordLandscape { get; set; }

    /// <summary>Optional header text on each Word page.</summary>
    public string? WordHeaderText { get; set; }

    /// <summary>Optional footer text on each Word page.</summary>
    public string? WordFooterText { get; set; }

    /// <summary>Optional fixed Word table cell width in twips.</summary>
    public int? WordColumnWidthTwips { get; set; }
}

/// <summary>Result of a single workbook conversion in a batch.</summary>
public sealed record ExcelConversionResult(string InputFile, string? OutputFile, bool Success, string? Error);
