using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Ss = DocumentFormat.OpenXml.Spreadsheet;
using TerraPDF.Core;
using TerraPDF.Helpers;
using TerraPDF.Infra;

namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    /// <summary>
    /// Worksheet print settings are opt-in. A print area is rendered as a series
    /// of independent, paginated PDF tables, split at explicit Excel page breaks.
    /// This is a bounded approximation, not a replica of Excel's print engine.
    /// </summary>
    private static Task WriteWorksheetPrintPdfAsync(
        string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        var layout = ReadWorksheetPrintLayout(inputFile, options, cancellationToken);
        return Task.Run(() =>
        {
            TerraPDF.Core.Document.Create(document =>
            {
                foreach (var chunk in layout.Chunks)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.Page(page =>
                    {
                        page.Size(layout.Landscape
                            ? PageSize.Landscape(PageSize.A4)
                            : PageSize.A4);
                        page.Margin(top: layout.TopMarginPoints,
                            right: layout.RightMarginPoints,
                            bottom: layout.BottomMarginPoints,
                            left: layout.LeftMarginPoints);
                        page.DefaultTextStyle(style => style.FontSize(9));

                        if (!string.IsNullOrWhiteSpace(options.DocumentTitle))
                            page.Header().Text(options.DocumentTitle).Bold().FontSize(14);
                        page.Footer().AlignCenter().Text(t =>
                        {
                            t.Span("Page ");
                            t.CurrentPageNumber();
                            t.Span(" / ");
                            t.TotalPages();
                        });

                        page.Content().Table(table =>
                        {
                            table.ColumnsDefinition(cols =>
                            {
                                foreach (var width in chunk.ColumnWidths)
                                    cols.RelativeColumn(width);
                            });
                            void RenderHeader(PdfStyledRow heading) =>
                                table.HeaderRow(row =>
                                {
                                    foreach (var cell in heading.Cells)
                                        RenderPdfCell(row.Cell(columnSpan: cell.ColumnSpan),
                                            ScalePrintCell(cell, layout.Scale), true);
                                });
                            RenderHeader(chunk.Header);
                            if (chunk.AdditionalHeaders is not null)
                                foreach (var heading in chunk.AdditionalHeaders)
                                    RenderHeader(heading);
                            foreach (var line in chunk.Rows)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                table.Row(row =>
                                {
                                    foreach (var cell in line.Cells)
                                        RenderPdfCell(
                                            row.Cell(columnSpan: cell.ColumnSpan,
                                                rowSpan: cell.RowSpan),
                                            ScalePrintCell(cell, layout.Scale), false);
                                });
                            }
                        });
                    });
                }
            }).PublishPdf(outputFile);
            cancellationToken.ThrowIfCancellationRequested();
        }, cancellationToken);
    }

    private static PdfStyledCell ScalePrintCell(PdfStyledCell cell, double scale)
    {
        if (Math.Abs(scale - 1) < 0.0001) return cell;
        return cell with
        {
            FontSize = Math.Clamp(cell.FontSize * scale, 4, 48),
            VerticalPadding = Math.Clamp(cell.VerticalPadding * scale, 1, 80),
            Picture = cell.Picture is null ? null : cell.Picture with
            {
                WidthPoints = Math.Clamp(cell.Picture.WidthPoints * scale, 5, 250)
            },
            TopBorder = ScaleBorder(cell.TopBorder, scale),
            BottomBorder = ScaleBorder(cell.BottomBorder, scale),
            LeftBorder = ScaleBorder(cell.LeftBorder, scale),
            RightBorder = ScaleBorder(cell.RightBorder, scale)
        };
    }

    private static PdfCellBorder? ScaleBorder(PdfCellBorder? border, double scale) =>
        border is null ? null : border with { Width = Math.Clamp(border.Width * scale, 0.2, 8) };

    internal sealed record PdfPrintLayout(
        IReadOnlyList<PdfFormattingSnapshot> Chunks,
        bool Landscape, double LeftMarginPoints, double RightMarginPoints,
        double TopMarginPoints, double BottomMarginPoints, double Scale);

    internal static PdfPrintLayout ReadWorksheetPrintLayout(
        string filePath, ExcelConversionOptions options, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(options);
        if (!options.PdfPreserveCellFormatting)
            throw new ArgumentException(
                "PdfUseWorksheetPrintSettings requires PdfPreserveCellFormatting.",
                nameof(options));
        if (options.PdfPrintScalePercent is < 10 or > 400)
            throw new ArgumentOutOfRangeException(nameof(options.PdfPrintScalePercent),
                "Print scale must be between 10% and 400%.");

        ValidateWorkbookInput(filePath, options, cancellationToken);
        using var workbook = new XLWorkbook(filePath);
        cancellationToken.ThrowIfCancellationRequested();
        var worksheet = options.WorksheetName is not null
            ? workbook.Worksheets.FirstOrDefault(w =>
                string.Equals(w.Name, options.WorksheetName, StringComparison.OrdinalIgnoreCase))
            : workbook.Worksheets.Skip(options.WorksheetIndex).FirstOrDefault();
        if (worksheet is null)
            throw new ArgumentException("Selected worksheet does not exist.", nameof(options));

        var settings = worksheet.PageSetup;
        var regions = (string.IsNullOrWhiteSpace(options.RangeAddress)
                ? settings.PrintAreas.Select(x => x.RangeAddress).ToList()
                : new List<IXLRangeAddress> { worksheet.Range(options.RangeAddress).RangeAddress })
            .ToList();
        if (regions.Count == 0)
        {
            var used = worksheet.RangeUsed()
                ?? throw new InvalidDataException("Worksheet has no printable content.");
            regions.Add(used.RangeAddress);
        }

        // PDF tables cannot reproduce Excel's fit-to-page algorithm. Reject
        // its workbook settings rather than silently treating them as percent.
        var scalePercent = options.PdfPrintScalePercent ?? settings.Scale;
        if (scalePercent is < 10 or > 400)
            throw new NotSupportedException("Worksheet print scale must be 10..400%.");

        var horizontal = new SortedSet<int>();
        var vertical = new SortedSet<int>();
        using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
        {
            var wb = spreadsheet.WorkbookPart
                ?? throw new InvalidDataException("XLSX workbook part is missing.");
            var workbookXml = wb.Workbook
                ?? throw new InvalidDataException("XLSX workbook XML is missing.");
            var sheet = (workbookXml.Sheets
                ?? throw new InvalidDataException("XLSX worksheet metadata is missing."))
                .Elements<Ss.Sheet>()
                .FirstOrDefault(s =>
                    string.Equals(s.Name?.Value, worksheet.Name, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidDataException("Worksheet metadata was not found.");
            var relationshipId = sheet.Id?.Value
                ?? throw new InvalidDataException("XLSX worksheet relationship is missing.");
            if (wb.GetPartById(relationshipId) is not WorksheetPart part)
                throw new InvalidDataException("XLSX worksheet part is missing.");
            var worksheetXml = part.Worksheet
                ?? throw new InvalidDataException("XLSX worksheet XML is missing.");
            if (options.PdfPrintScalePercent is null &&
                worksheetXml.GetFirstChild<Ss.SheetProperties>()?
                    .PageSetupProperties?.FitToPage?.Value == true)
                throw new NotSupportedException(
                    "Excel fit-to-N-pages cannot be reproduced by the PDF table renderer. " +
                    "Set PdfPrintScalePercent to use an explicit scaling percentage.");
            var h = worksheetXml.Elements<Ss.RowBreaks>().FirstOrDefault();
            var v = worksheetXml.Elements<Ss.ColumnBreaks>().FirstOrDefault();
            if (h is not null)
                foreach (var b in h.Elements<Ss.Break>())
                    if (b.ManualPageBreak?.Value == true && b.Id?.Value is uint index && index > 0)
                        horizontal.Add(checked((int)index));
            if (v is not null)
                foreach (var b in v.Elements<Ss.Break>())
                    if (b.ManualPageBreak?.Value == true && b.Id?.Value is uint index && index > 0)
                        vertical.Add(checked((int)index));
        }

        var chunks = new List<PdfFormattingSnapshot>();
        foreach (var region in regions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int r1 = region.FirstAddress.RowNumber, r2 = region.LastAddress.RowNumber;
            int c1 = region.FirstAddress.ColumnNumber, c2 = region.LastAddress.ColumnNumber;
            ValidateSelectedCells(r1, r2, c1, c2, options);

            // Obtain a stable repeating header from the region's first row,
            // or the workbook's dedicated print-title row, if provided.
            var titleRow = settings.FirstRowToRepeatAtTop > 0
                ? settings.FirstRowToRepeatAtTop : r1;
            var titleLastRow = settings.LastRowToRepeatAtTop > 0
                ? settings.LastRowToRepeatAtTop : titleRow;
            if (titleRow > r2 || titleRow < 1 || titleLastRow < titleRow)
                throw new NotSupportedException("Print title row is outside the print area.");
            if (titleRow > r1)
                throw new NotSupportedException(
                    "Print title row must precede or equal the first printed row.");

            var rowStart = Math.Max(r1, titleLastRow + 1);
            var rowSegments = SplitAtBreaks(rowStart, r2, horizontal);
            var colSegments = SplitAtBreaks(c1, c2, vertical);
            foreach (var (cs, ce) in colSegments)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var titleSnapshot = ReadPdfFormatting(filePath, new ExcelConversionOptions
                {
                    WorksheetName = worksheet.Name,
                    RangeAddress = worksheet.Range(titleRow, cs, titleLastRow, ce)
                        .RangeAddress.ToString(),
                    PdfPreserveCellFormatting = true,
                    PdfIncludeEmbeddedImages = options.PdfIncludeEmbeddedImages,
                    PdfMaxEmbeddedImageBytes = options.PdfMaxEmbeddedImageBytes,
                    PdfMaxEmbeddedImages = options.PdfMaxEmbeddedImages,
                PdfIgnoreImagesOutsideSelectedRange = true
                }, cancellationToken);
                var colTitle = titleSnapshot.Header;
                var extraTitleRows = titleSnapshot.Rows;
                foreach (var (rs, re) in rowSegments)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    // A range's first row is treated as a header by the shared
                    // extractor. Here it is the first data row of this slice.
                    PdfFormattingSnapshot page;
                    if (rs <= re)
                    {
                        page = ReadPdfFormatting(filePath, new ExcelConversionOptions
                        {
                            WorksheetName = worksheet.Name,
                            RangeAddress = worksheet.Range(rs, cs, re, ce)
                                .RangeAddress.ToString(),
                            PdfPreserveCellFormatting = true,
                            PdfIncludeEmbeddedImages = options.PdfIncludeEmbeddedImages,
                            PdfMaxEmbeddedImageBytes = options.PdfMaxEmbeddedImageBytes,
                            PdfMaxEmbeddedImages = options.PdfMaxEmbeddedImages,
                PdfIgnoreImagesOutsideSelectedRange = true,
                            SkipEmptyRows = options.SkipEmptyRows
                        }, cancellationToken);
                        chunks.Add(new PdfFormattingSnapshot(page.ColumnWidths, colTitle,
                            new[] { page.Header }.Concat(page.Rows).ToList(),
                            extraTitleRows));
                    }
                    else
                    {
                        var widths = Enumerable.Range(cs, ce - cs + 1)
                            .Select(i => (double)Math.Clamp(worksheet.Column(i).Width, 3, 80))
                            .ToList();
                        chunks.Add(new PdfFormattingSnapshot(widths, colTitle,
                            Array.Empty<PdfStyledRow>(), extraTitleRows));
                    }
                }
            }
        }
        if (chunks.Count > 128)
            throw new NotSupportedException(
                "More than 128 explicit PDF print sections are not supported.");

        static double Points(double inches)
        {
            if (double.IsNaN(inches) || double.IsInfinity(inches) || inches < 0.1 || inches > 3)
                throw new ArgumentOutOfRangeException(nameof(inches),
                    "PDF print margins must be between 0.1 and 3 inches.");
            return inches * 72;
        }
        return new PdfPrintLayout(chunks,
            settings.PageOrientation == XLPageOrientation.Landscape,
            Points(settings.Margins.Left), Points(settings.Margins.Right),
            Points(settings.Margins.Top), Points(settings.Margins.Bottom),
            scalePercent / 100.0);
    }

    private static IReadOnlyList<(int Start, int End)> SplitAtBreaks(
        int start, int end, IEnumerable<int> breaks)
    {
        if (end < start) return new[] { (start, end) };
        var output = new List<(int Start, int End)>();
        foreach (var boundary in breaks.Where(x => x >= start && x < end).Distinct().Order())
        {
            output.Add((start, boundary));
            start = boundary + 1;
        }
        output.Add((start, end));
        return output;
    }
}
