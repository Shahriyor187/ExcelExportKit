using ClosedXML.Excel;
using System.Globalization;
using TerraPDF.Core;
using TerraPDF.Helpers;

namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    /// <summary>
    /// Creates a paginated PDF while preserving supported cell formatting:
    /// RGB fills and text colors, bold/italic text, horizontal alignment and
    /// proportional column widths. This is a formatting-aware table renderer,
    /// not a pixel-exact Excel print engine; drawings, formulas, merged cells,
    /// page breaks and print area are not preserved in this first version.
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
                        table.HeaderRow(row =>
                        {
                            foreach (var cell in snapshot.Header)
                            {
                                var content = row.Cell();
                                content = content.Background(cell.Background ?? "#1a4a8a");
                                if (cell.AlignRight) content = content.AlignRight();
                                else if (cell.AlignCenter) content = content.AlignCenter();
                                var text = content.Padding(5).Text(cell.Text);
                                if (cell.Bold || cell.Background is null) text.Bold();
                                if (cell.Italic) text.Italic();
                                text.FontColor(cell.Foreground ?? (cell.Background is null ? "#FFFFFF" : "#111111"));
                            }
                        });

                        foreach (var cells in snapshot.Rows)
                        {
                            table.Row(row =>
                            {
                                foreach (var cell in cells)
                                {
                                    var content = row.Cell();
                                    if (cell.Background is not null)
                                        content = content.Background(cell.Background);
                                    if (cell.AlignRight) content = content.AlignRight();
                                    else if (cell.AlignCenter) content = content.AlignCenter();
                                    var text = content.BorderBottom(0.5, "#dddddd")
                                        .Padding(5).Text(cell.Text);
                                    if (cell.Bold) text.Bold();
                                    if (cell.Italic) text.Italic();
                                    if (cell.Foreground is not null)
                                        text.FontColor(cell.Foreground);
                                }
                            });
                        }
                    });
                });
            }).PublishPdf(outputFile);
        });
    }

    private static PdfFormattingSnapshot ReadPdfFormatting(string filePath, ExcelConversionOptions options)
    {
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

        var used = worksheet.RangeUsed()
            ?? throw new InvalidDataException("The selected worksheet does not contain any data.");
        var firstRow = used.RangeAddress.FirstAddress.RowNumber;
        var lastRow = used.RangeAddress.LastAddress.RowNumber;
        var firstCol = used.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = used.RangeAddress.LastAddress.ColumnNumber;

        var widths = new List<double>();
        for (var col = firstCol; col <= lastCol; col++)
            widths.Add(Math.Clamp(worksheet.Column(col).Width, 3, 80));

        IReadOnlyList<PdfStyledCell> GetRow(int rowIndex)
        {
            var result = new List<PdfStyledCell>();
            for (var col = firstCol; col <= lastCol; col++)
            {
                var cell = worksheet.Cell(rowIndex, col);
                var alignment = cell.Style.Alignment.Horizontal;
                result.Add(new PdfStyledCell(
                    cell.GetFormattedString(CultureInfo.InvariantCulture),
                    ToPdfRgb(cell.Style.Fill.BackgroundColor),
                    ToPdfRgb(cell.Style.Font.FontColor),
                    cell.Style.Font.Bold,
                    cell.Style.Font.Italic,
                    alignment == XLAlignmentHorizontalValues.Right,
                    alignment == XLAlignmentHorizontalValues.Center));
            }
            return result;
        }

        var header = GetRow(firstRow);
        var rows = new List<IReadOnlyList<PdfStyledCell>>();
        for (var row = firstRow + 1; row <= lastRow; row++)
            rows.Add(GetRow(row));
        return new PdfFormattingSnapshot(widths, header, rows);
    }

    private static string? ToPdfRgb(XLColor color)
    {
        if (color.ColorType != XLColorType.Color)
            return null;
        var rgb = color.Color;
        return $"#{rgb.R:X2}{rgb.G:X2}{rgb.B:X2}";
    }

    private sealed record PdfStyledCell(
        string Text, string? Background, string? Foreground,
        bool Bold, bool Italic, bool AlignRight, bool AlignCenter);

    private sealed record PdfFormattingSnapshot(
        IReadOnlyList<double> ColumnWidths,
        IReadOnlyList<PdfStyledCell> Header,
        IReadOnlyList<IReadOnlyList<PdfStyledCell>> Rows);
}
