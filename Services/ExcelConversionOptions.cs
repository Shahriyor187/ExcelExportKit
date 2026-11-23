namespace ExportFiles.Services;

/// <summary>Settings for worksheet selection and conversion output.</summary>
public sealed class ExcelConversionOptions
{
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
