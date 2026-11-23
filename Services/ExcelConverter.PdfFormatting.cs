using ClosedXML.Excel;
using System.Globalization;
using TerraPDF.Core;
using TerraPDF.Helpers;
using TerraPDF.Infra;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("ExcelExportKit.Tests")]

namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    /// <summary>
    /// Renders a styled, paginated table from a selected Excel worksheet, including fully
    /// selected merged cells, basic borders, source row heights and font sizes.
    /// This remains a table renderer rather than an exact Excel print-layout engine.
    /// </summary>
    private static Task WriteFormattingAwarePdfAsync(
        string inputFile, string outputFile, ExcelConversionOptions options)
    {
        var snapshot = ReadPdfFormatting(inputFile, options);
        return Task.Run(() =>
        {
            TerraPDF.Core.Document.Create(document =>
            {
                document.Page(page =>
                {
                    page.Size(options.PdfLandscape ? PageSize.Landscape(PageSize.A4) : PageSize.A4);
                    page.Margin(1.5, Unit.Centimetre);
                    page.DefaultTextStyle(style => style.FontSize(9));

                    if (!string.IsNullOrWhiteSpace(options.DocumentTitle))
                        page.Header().Text(options.DocumentTitle).Bold().FontSize(14);

                    page.Footer().AlignCenter().Text(text =>
                    {
                        text.Span("Page ");
                        text.CurrentPageNumber();
                        text.Span(" / ");
                        text.TotalPages();
                    });

                    page.Content().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            foreach (var width in snapshot.ColumnWidths)
                                columns.RelativeColumn(width);
                        });

                        // HeaderRow repeats on subsequent pages. Spans within the header
                        // are allowed, but spans crossing into data rows are rejected
                        // during extraction to prevent corrupt repeated headers.
                        table.HeaderRow(row =>
                        {
                            foreach (var cell in snapshot.Header.Cells)
                                RenderPdfCell(row.Cell(columnSpan: cell.ColumnSpan), cell, header: true);
                        });

                        foreach (var values in snapshot.Rows)
                        {
                            table.Row(row =>
                            {
                                foreach (var cell in values.Cells)
                                {
                                    RenderPdfCell(
                                        row.Cell(columnSpan: cell.ColumnSpan, rowSpan: cell.RowSpan),
                                        cell, header: false);
                                }
                            });
                        }
                    });
                });
            }).PublishPdf(outputFile);
        });
    }

    private static void RenderPdfCell(IContainer container, PdfStyledCell cell, bool header)
    {
        var current = container;
        if (cell.Background is not null || header)
            current = current.Background(cell.Background ?? "#1a4a8a");

        if (cell.TopBorder is not null)
            current = current.BorderTop(cell.TopBorder.Width, cell.TopBorder.Color);
        if (cell.BottomBorder is not null)
            current = current.BorderBottom(cell.BottomBorder.Width, cell.BottomBorder.Color);
        else if (!header)
            current = current.BorderBottom(0.5, "#dddddd");
        if (cell.LeftBorder is not null)
            current = current.BorderLeft(cell.LeftBorder.Width, cell.LeftBorder.Color);
        if (cell.RightBorder is not null)
            current = current.BorderRight(cell.RightBorder.Width, cell.RightBorder.Color);

        current = current.PaddingVertical(cell.VerticalPadding).PaddingHorizontal(5);
        if (cell.AlignRight)
            current = current.AlignRight();
        else if (cell.AlignCenter)
            current = current.AlignCenter();

        if (cell.Picture is not null)
        {
            current.Column(column =>
            {
                if (!string.IsNullOrWhiteSpace(cell.Text))
                {
                    var caption = column.Item().Text(cell.Text);
                    if (header || cell.Bold) caption.Bold();
                    if (cell.Italic) caption.Italic();
                    caption.FontSize(cell.FontSize);
                    caption.FontColor(cell.Foreground ??
                        (header && cell.Background is null ? "#FFFFFF" : "#111111"));
                }
                // This image is intentionally IN the cell's flow layout; it is
                // not a pixel-exact recreation of Excel's floating drawings.
                column.Item().AlignCenter().Image(cell.Picture.Bytes, cell.Picture.WidthPoints);
            });
        }
        else
        {
            var text = current.Text(cell.Text);
            if (header || cell.Bold) text.Bold();
            if (cell.Italic) text.Italic();
            text.FontSize(cell.FontSize);
            text.FontColor(cell.Foreground ?? (header && cell.Background is null ? "#FFFFFF" : "#111111"));
        }
    }

    /// <summary>
    /// Internal to permit structural tests for merge shapes, boundaries and styles.
    /// No PDF-related API is exposed to consumers in this iteration.
    /// </summary>
    internal static PdfFormattingSnapshot ReadPdfFormatting(string filePath, ExcelConversionOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath))
            throw new FileNotFoundException("The Excel input file was not found.", filePath);
        if (options.WorksheetIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(options.WorksheetIndex));

        using var workbook = new XLWorkbook(filePath);
        var worksheet = options.WorksheetName is not null
            ? workbook.Worksheets.FirstOrDefault(w =>
                string.Equals(w.Name, options.WorksheetName, StringComparison.OrdinalIgnoreCase))
            : workbook.Worksheets.Skip(options.WorksheetIndex).FirstOrDefault();
        if (worksheet is null)
            throw new ArgumentException("The selected worksheet does not exist.", nameof(options));

        var used = string.IsNullOrWhiteSpace(options.RangeAddress)
            ? worksheet.RangeUsed()
            : worksheet.Range(options.RangeAddress);
        if (used is null)
            throw new InvalidDataException("The selected worksheet range does not contain any data.");

        var firstRow = used.RangeAddress.FirstAddress.RowNumber;
        var lastRow = used.RangeAddress.LastAddress.RowNumber;
        var firstCol = used.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = used.RangeAddress.LastAddress.ColumnNumber;

        // A merge is allowed only when its entire rectangle is selected.
        // Cropped spans or a vertical merge from the repeating header into
        // data would silently distort the exported report.
        var anchors = new Dictionary<(int Row, int Column), (int Rows, int Columns)>();
        var occupied = new HashSet<(int Row, int Column)>();
        foreach (var merged in worksheet.MergedRanges)
        {
            var a = merged.RangeAddress.FirstAddress;
            var b = merged.RangeAddress.LastAddress;
            if (b.RowNumber < firstRow || a.RowNumber > lastRow ||
                b.ColumnNumber < firstCol || a.ColumnNumber > lastCol)
                continue;

            if (a.RowNumber < firstRow || b.RowNumber > lastRow ||
                a.ColumnNumber < firstCol || b.ColumnNumber > lastCol)
                throw new NotSupportedException(
                    $"Merged range {merged.RangeAddress} crosses the selected PDF range. Select its entire rectangle.");

            if (a.RowNumber == firstRow && b.RowNumber > firstRow)
                throw new NotSupportedException(
                    $"Merged range {merged.RangeAddress} crosses the repeated PDF header. " +
                    "Use a selection whose first row is not vertically merged with the data.");

            anchors.Add((a.RowNumber, a.ColumnNumber),
                (b.RowNumber - a.RowNumber + 1, b.ColumnNumber - a.ColumnNumber + 1));
            for (var row = a.RowNumber; row <= b.RowNumber; row++)
                for (var col = a.ColumnNumber; col <= b.ColumnNumber; col++)
                    if (row != a.RowNumber || col != a.ColumnNumber)
                        occupied.Add((row, col));
        }

        // Validate merge boundaries before checking for empty cells: the selected
        // rectangle may include only covered (blank) cells of a merged range.
        if (used.IsEmpty())
            throw new InvalidDataException("The selected worksheet range does not contain any data.");

        // Images are opt-in because pictures are floating drawings in Excel.
        // This renderer supports only one PNG/JPEG anchored inside each selected,
        // unmerged cell. The picture is placed below any existing cell text.
        var pictures = new Dictionary<(int Row, int Column), PdfCellPicture>();
        if (options.PdfIncludeEmbeddedImages)
        {
            if (!options.PdfPreserveCellFormatting)
                throw new ArgumentException("Enable PdfPreserveCellFormatting for embedded PDF images.", nameof(options));
            if (options.PdfMaxEmbeddedImages is < 0 or > 256)
                throw new ArgumentOutOfRangeException(nameof(options.PdfMaxEmbeddedImages));
            if (options.PdfMaxEmbeddedImageBytes is < 1024 or > 16 * 1024 * 1024)
                throw new ArgumentOutOfRangeException(nameof(options.PdfMaxEmbeddedImageBytes));

            var count = 0;
            foreach (var picture in worksheet.Pictures)
            {
                count++;
                if (count > options.PdfMaxEmbeddedImages)
                    throw new InvalidDataException(
                        $"Workbook contains more than {options.PdfMaxEmbeddedImages} pictures.");

                // Free-floating pictures have no meaningful cell anchor.
                if (picture.Placement.ToString() == "FreeFloating")
                    throw new NotSupportedException(
                        $"Picture '{picture.Name}' is free-floating; only cell-anchored pictures are supported.");
                var anchorCell = picture.TopLeftCell;
                var row = anchorCell.Address.RowNumber;
                var col = anchorCell.Address.ColumnNumber;
                if (row < firstRow || row > lastRow || col < firstCol || col > lastCol)
                    throw new NotSupportedException(
                        $"Picture '{picture.Name}' is outside the selected PDF cell range.");
                if (occupied.Contains((row, col)) || anchors.ContainsKey((row, col)))
                    throw new NotSupportedException(
                        $"Picture '{picture.Name}' overlaps a merged cell, which is not supported.");
                if (picture.Left != 0 || picture.Top != 0 ||
                    picture.GetOffset(ClosedXML.Excel.Drawings.XLMarkerPosition.TopLeft).X != 0 ||
                    picture.GetOffset(ClosedXML.Excel.Drawings.XLMarkerPosition.TopLeft).Y != 0)
                    throw new NotSupportedException(
                        $"Picture '{picture.Name}' uses an offset; position it at a cell origin.");
                if (!pictures.TryAdd((row, col), null!))
                    throw new NotSupportedException(
                        $"Multiple pictures in cell {anchorCell.Address} are not supported.");

                var format = picture.Format.ToString();
                if (!string.Equals(format, "Png", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(format, "Jpeg", StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException(
                        $"Picture '{picture.Name}' format {format} is unsupported. Use PNG or JPEG.");
                if (picture.ImageStream.Length > options.PdfMaxEmbeddedImageBytes)
                    throw new InvalidDataException(
                        $"Picture '{picture.Name}' exceeds the maximum byte size.");
                if (picture.Width is < 1 or > 4096 || picture.Height is < 1 or > 4096)
                    throw new InvalidDataException(
                        $"Picture '{picture.Name}' has unsupported dimensions.");

                var bytes = picture.ImageStream.ToArray();
                // Limit image width to fit the cell in the PDF. 96 DPI source
                // pixels convert to 72-point PDF units, capped for narrow cells.
                var targetWidth = Math.Clamp(picture.Width * 0.75, 20, 110);
                pictures[(row, col)] = new PdfCellPicture(bytes, targetWidth);
            }
        }

        var widths = new List<double>();
        for (var col = firstCol; col <= lastCol; col++)
            widths.Add(Math.Clamp(worksheet.Column(col).Width, 3, 80));

        PdfStyledRow GetRow(int rowIndex)
        {
            var result = new List<PdfStyledCell>();
            var containsMergedCells = false;
            for (var col = firstCol; col <= lastCol; col++)
            {
                if (occupied.Contains((rowIndex, col)))
                {
                    containsMergedCells = true;
                    continue;
                }

                var span = anchors.TryGetValue((rowIndex, col), out var rect)
                    ? rect : (Rows: 1, Columns: 1);
                if (span.Rows > 1 || span.Columns > 1)
                    containsMergedCells = true;

                var cell = worksheet.Cell(rowIndex, col);
                var border = cell.Style.Border;
                var rightEdge = worksheet.Cell(rowIndex, col + span.Columns - 1).Style.Border;
                var bottomEdge = worksheet.Cell(rowIndex + span.Rows - 1, col).Style.Border;
                var alignment = cell.Style.Alignment.Horizontal;
                var rowHeight = worksheet.Row(rowIndex).Height;
                // Excel height is measured in points; TerraPDF's cell height is
                // content-driven. Padding approximates a minimum visible row height.
                var padding = Math.Clamp((rowHeight - 11.0) / 2.0, 4, 30);

                result.Add(new PdfStyledCell(
                    cell.GetFormattedString(CultureInfo.InvariantCulture),
                    ToPdfRgb(cell.Style.Fill.BackgroundColor),
                    ToPdfRgb(cell.Style.Font.FontColor),
                    cell.Style.Font.Bold,
                    cell.Style.Font.Italic,
                    alignment == XLAlignmentHorizontalValues.Right,
                    alignment == XLAlignmentHorizontalValues.Center,
                    Math.Clamp(cell.Style.Font.FontSize, 6, 24),
                    padding,
                    span.Columns,
                    span.Rows,
                    AsPdfBorder(border.TopBorder, border.TopBorderColor),
                    AsPdfBorder(bottomEdge.BottomBorder, bottomEdge.BottomBorderColor),
                    AsPdfBorder(border.LeftBorder, border.LeftBorderColor),
                    AsPdfBorder(rightEdge.RightBorder, rightEdge.RightBorderColor),
                    pictures.GetValueOrDefault((rowIndex, col))));
            }
            return new PdfStyledRow(result, containsMergedCells);
        }

        var header = GetRow(firstRow);
        var rows = new List<PdfStyledRow>();
        for (var row = firstRow + 1; row <= lastRow; row++)
        {
            var styled = GetRow(row);
            // Never remove rows occupied by a rowspan: that would cause the
            // TerraPDF table to put subsequent cells in the wrong positions.
            if (options.SkipEmptyRows &&
                styled.Cells.All(c => string.IsNullOrWhiteSpace(c.Text)) &&
                !styled.ContainsMergedCells)
                continue;
            rows.Add(styled);
        }

        return new PdfFormattingSnapshot(widths, header, rows);
    }

    private static PdfCellBorder? AsPdfBorder(XLBorderStyleValues style, XLColor color)
    {
        if (style == XLBorderStyleValues.None)
            return null;

        // PDF currently supports solid per-edge strokes; Excel's dash patterns
        // and double-line borders are intentionally approximated.
        var width = style switch
        {
            XLBorderStyleValues.Hair => 0.35,
            XLBorderStyleValues.Thin => 0.7,
            XLBorderStyleValues.Medium => 1.25,
            XLBorderStyleValues.Thick => 2.25,
            _ => 0.8
        };
        return new PdfCellBorder(width, ToPdfRgb(color) ?? "#505050");
    }

    private static string? ToPdfRgb(XLColor color)
    {
        if (color.ColorType != XLColorType.Color)
            return null;
        var rgb = color.Color;
        return $"#{rgb.R:X2}{rgb.G:X2}{rgb.B:X2}";
    }

    internal sealed record PdfCellBorder(double Width, string Color);

    internal sealed record PdfCellPicture(byte[] Bytes, double WidthPoints);

    internal sealed record PdfStyledCell(
        string Text, string? Background, string? Foreground,
        bool Bold, bool Italic, bool AlignRight, bool AlignCenter,
        double FontSize, double VerticalPadding, int ColumnSpan, int RowSpan,
        PdfCellBorder? TopBorder, PdfCellBorder? BottomBorder,
        PdfCellBorder? LeftBorder, PdfCellBorder? RightBorder,
        PdfCellPicture? Picture);

    internal sealed record PdfStyledRow(
        IReadOnlyList<PdfStyledCell> Cells, bool ContainsMergedCells);

    internal sealed record PdfFormattingSnapshot(
        IReadOnlyList<double> ColumnWidths,
        PdfStyledRow Header,
        IReadOnlyList<PdfStyledRow> Rows);
}
