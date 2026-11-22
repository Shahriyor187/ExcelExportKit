namespace ExportFiles.Services;

/// <summary>Settings for worksheet selection and conversion output.</summary>
public sealed class ExcelConversionOptions
{
    /// <summary>Select a worksheet by name. Takes precedence over WorksheetIndex.</summary>
    public string? WorksheetName { get; set; }

    /// <summary>Zero-based worksheet index; defaults to the first worksheet.</summary>
    public int WorksheetIndex { get; set; }

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

    /// <summary>Word table style identifier.</summary>
    public string WordTableStyle { get; set; } = "TableGrid";

    /// <summary>
    /// Preserve supported Excel fill/font colors, bold text, horizontal alignment and column-width proportions
    /// when exporting to PDF. Off by default to preserve the existing PDF appearance.
    /// Merged cells, images, borders and full Excel print-layout fidelity are not yet supported.
    /// </summary>
    public bool PdfPreserveCellFormatting { get; set; }

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
