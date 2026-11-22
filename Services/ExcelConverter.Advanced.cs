using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.Text;
using System.Text.Json;
using TerraPDF.Core;
using TerraPDF.Helpers;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace ExportFiles.Services;

public static partial class ExcelConverter
{
    /// <summary>Converts the selected worksheet to CSV.</summary>
    public static async Task ToCSVAsync(string inputFile, string outputFile, ExcelConversionOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        ValidateCsvDelimiter(options.CsvDelimiter);
        var sheet = ReadWorksheet(inputFile, options);
        var output = new StringBuilder();
        AppendSelectedCsvRow(output, sheet.Headers.Select(heading => SafeCsvValue(heading, options.CsvProtectFormulas)), options.CsvDelimiter);
        foreach (var row in sheet.Rows)
            AppendSelectedCsvRow(output, row.Select(value => SafeCsvValue(value, options.CsvProtectFormulas)), options.CsvDelimiter);
        await File.WriteAllTextAsync(outputFile, output.ToString(), Utf8WithoutBom).ConfigureAwait(false);
    }

    /// <summary>Converts the selected worksheet to JSON with optional native numbers and booleans.</summary>
    public static async Task ToJsonAsync(string inputFile, string outputFile, ExcelConversionOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        var sheet = ReadWorksheet(inputFile, options);
        var records = new List<Dictionary<string, object?>>();
        IEnumerable<IReadOnlyList<object?>> dataRows = options.PreserveJsonTypes
            ? sheet.TypedRows
            : sheet.Rows.Select(r => (IReadOnlyList<object?>)r.Cast<object?>().ToList());
        foreach (var row in dataRows)
        {
            var record = new Dictionary<string, object?>(StringComparer.Ordinal);
            for (var index = 0; index < sheet.Headers.Count; index++)
                record[sheet.Headers[index]] = row[index];
            records.Add(record);
        }
        await File.WriteAllTextAsync(outputFile, JsonSerializer.Serialize(records, JsonOptions), Utf8WithoutBom)
            .ConfigureAwait(false);
    }

    /// <summary>Exports the selected worksheet as a paginated PDF with repeated table headers.</summary>
    public static Task ToPDFAsync(string inputFile, string outputFile, ExcelConversionOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        var sheet = ReadWorksheet(inputFile, options);
        return Task.Run(() =>
        {
            TerraPDF.Core.Document.Create(document =>
            {
                document.Page(page =>
                {
                    page.Size(options.PdfLandscape ? PageSize.Landscape(PageSize.A4) : PageSize.A4);
                    page.Margin(1.5, Unit.Centimetre);
                    page.DefaultTextStyle(style => style.FontSize(9));
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
                            foreach (var _ in sheet.Headers)
                                columns.RelativeColumn();
                        });
                        table.HeaderRow(row =>
                        {
                            foreach (var heading in sheet.Headers)
                                row.Cell().Background("#1a4a8a").Padding(5)
                                    .Text(heading).Bold().FontColor(Color.White);
                        });
                        foreach (var values in sheet.Rows)
                        {
                            table.Row(row =>
                            {
                                foreach (var value in values)
                                    row.Cell().BorderBottom(0.5, "#dddddd").Padding(5).Text(value);
                            });
                        }
                    });
                });
            }).PublishPdf(outputFile);
        });
    }

    /// <summary>Converts the selected worksheet to a DOCX table with configurable title and style.</summary>
    public static Task ToWordAsync(string inputFile, string outputFile, ExcelConversionOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        var sheet = ReadWorksheet(inputFile, options);
        return Task.Run(() => WriteStyledWordDocument(sheet, outputFile, options));
    }

    /// <summary>Converts every worksheet to a separate file in the requested format.</summary>
    public static async Task<IReadOnlyList<ExcelConversionResult>> ConvertAllSheetsAsync(
        string inputFile, string outputDirectory, string format, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputFile);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        ArgumentException.ThrowIfNullOrWhiteSpace(format);
        var normalized = format.ToLowerInvariant();
        if (normalized is not ("csv" or "json" or "pdf" or "docx"))
            throw new ArgumentException("Format must be csv, json, pdf, or docx.", nameof(format));
        if (!File.Exists(inputFile))
            throw new FileNotFoundException("Workbook not found.", inputFile);

        using var workbook = new XLWorkbook(inputFile);
        var sheets = workbook.Worksheets.Select(w => w.Name).ToList();
        Directory.CreateDirectory(outputDirectory);
        var results = new List<ExcelConversionResult>(sheets.Count);

        foreach (var sheetName in sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var safeName = SanitizeFileName(sheetName);
            var destination = Path.Combine(outputDirectory, safeName + "." + normalized);
            var suffix = 2;
            while (File.Exists(destination) || results.Any(r => string.Equals(r.OutputFile, destination, StringComparison.OrdinalIgnoreCase)))
                destination = Path.Combine(outputDirectory, safeName + "_" + suffix++ + "." + normalized);

            var options = new ExcelConversionOptions { WorksheetName = sheetName };
            try
            {
                switch (normalized)
                {
                    case "csv": await ToCSVAsync(inputFile, destination, options).ConfigureAwait(false); break;
                    case "json": await ToJsonAsync(inputFile, destination, options).ConfigureAwait(false); break;
                    case "pdf": await ToPDFAsync(inputFile, destination, options).ConfigureAwait(false); break;
                    case "docx": await ToWordAsync(inputFile, destination, options).ConfigureAwait(false); break;
                }
                results.Add(new ExcelConversionResult(inputFile, destination, true, null));
            }
            catch (Exception error) when (error is not OperationCanceledException)
            {
                results.Add(new ExcelConversionResult(inputFile, null, false, $"{sheetName}: {error.Message}"));
            }
        }
        return results;
    }

    /// <summary>Converts several workbooks while reporting individual failures.</summary>
    public static async Task<IReadOnlyList<ExcelConversionResult>> ConvertBatchAsync(
        IEnumerable<string> inputFiles, string outputDirectory, string format, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(inputFiles);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        var normalized = format?.ToLowerInvariant() ?? throw new ArgumentNullException(nameof(format));
        if (normalized is not ("csv" or "json" or "pdf" or "docx"))
            throw new ArgumentException("Format must be csv, json, pdf, or docx.", nameof(format));
        Directory.CreateDirectory(outputDirectory);
        var results = new List<ExcelConversionResult>();
        foreach (var file in inputFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var baseName = SanitizeFileName(Path.GetFileNameWithoutExtension(file));
            var path = Path.Combine(outputDirectory, baseName + "." + normalized);
            var counter = 2;
            while (File.Exists(path) || results.Any(r => string.Equals(r.OutputFile, path, StringComparison.OrdinalIgnoreCase)))
                path = Path.Combine(outputDirectory, baseName + "_" + counter++ + "." + normalized);
            try
            {
                switch (normalized)
                {
                    case "csv": await ToCSVAsync(file, path).ConfigureAwait(false); break;
                    case "json": await ToJsonAsync(file, path).ConfigureAwait(false); break;
                    case "pdf": await ToPDFAsync(file, path).ConfigureAwait(false); break;
                    case "docx": await ToWordAsync(file, path).ConfigureAwait(false); break;
                }
                results.Add(new ExcelConversionResult(file, path, true, null));
            }
            catch (Exception error) when (error is not OperationCanceledException)
            {
                results.Add(new ExcelConversionResult(file, null, false, error.Message));
            }
        }
        return results;
    }

    private static SelectedWorksheet ReadWorksheet(string filePath, ExcelConversionOptions options)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath))
            throw new FileNotFoundException("The Excel input file was not found.", filePath);
        using var workbook = new XLWorkbook(filePath);
        if (workbook.Worksheets.Count == 0)
            throw new InvalidDataException("The workbook contains no worksheets.");
        if (options.WorksheetIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(options.WorksheetIndex));
        var worksheet = options.WorksheetName is not null
            ? workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, options.WorksheetName, StringComparison.OrdinalIgnoreCase))
            : workbook.Worksheets.Skip(options.WorksheetIndex).FirstOrDefault();
        if (worksheet is null)
            throw new ArgumentException("The selected worksheet does not exist.", nameof(options));
        var usedRange = worksheet.RangeUsed()
            ?? throw new InvalidDataException("The selected worksheet does not contain any data.");
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
        var firstColumn = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumn = usedRange.RangeAddress.LastAddress.ColumnNumber;
        var rawHeaders = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1)
            .Select(col => worksheet.Cell(firstRow, col).GetFormattedString(CultureInfo.InvariantCulture).Trim())
            .ToList();
        var headers = MakeUniqueHeaders(rawHeaders);
        var displayRows = new List<IReadOnlyList<string>>();
        var typedRows = new List<IReadOnlyList<object?>>();

        for (var row = firstRow + 1; row <= lastRow; row++)
        {
            var display = new List<string>();
            var typed = new List<object?>();
            for (var col = firstColumn; col <= lastColumn; col++)
            {
                var cell = worksheet.Cell(row, col);
                display.Add(cell.GetFormattedString(CultureInfo.InvariantCulture));
                typed.Add(cell.DataType switch
                {
                    XLDataType.Number => cell.GetDouble(),
                    XLDataType.Boolean => cell.GetBoolean(),
                    XLDataType.DateTime => cell.GetDateTime(),
                    XLDataType.TimeSpan => cell.GetTimeSpan().ToString(),
                    XLDataType.Blank => null,
                    _ => cell.GetFormattedString(CultureInfo.InvariantCulture)
                });
            }
            displayRows.Add(display);
            typedRows.Add(typed);
        }

        if (headers.Count == 0)
            throw new InvalidDataException("The worksheet has no columns.");
        return new SelectedWorksheet(headers, displayRows, typedRows);
    }

    private static string SafeCsvValue(string input, bool protect)
    {
        return protect && input.TrimStart().Length > 0 &&
               input.TrimStart()[0] is '=' or '+' or '-' or '@'
            ? "'" + input
            : input;
    }

    private static void ValidateCsvDelimiter(char delimiter)
    {
        if (delimiter is '\0' or '\r' or '\n' or '"')
            throw new ArgumentException("Delimiter cannot be a newline or quote.", nameof(delimiter));
    }

    private static void AppendSelectedCsvRow(StringBuilder builder, IEnumerable<string> values, char delimiter)
    {
        builder.AppendJoin(delimiter, values.Select(value =>
            value.IndexOfAny([delimiter, '"', '\r', '\n']) < 0
                ? value
                : "\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\""));
        builder.AppendLine();
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var result = new string(name.Select(c => invalid.Contains(c) || c is '/' or '\\' ? '_' : c).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(result) ? "worksheet" : result;
    }

    private static void WriteStyledWordDocument(SelectedWorksheet sheet, string outputPath, ExcelConversionOptions options)
    {
        using var word = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Wp.Document(new Wp.Body());
        var body = mainPart.Document.Body!;
        if (!string.IsNullOrWhiteSpace(options.DocumentTitle))
        {
            var title = new Wp.Paragraph(new Wp.Run(new Wp.RunProperties(new Wp.Bold()),
                new Wp.Text(options.DocumentTitle)));
            body.Append(title);
        }
        var table = new Wp.Table();
        if (!string.IsNullOrWhiteSpace(options.WordTableStyle))
            table.Append(new Wp.TableProperties(new Wp.TableStyle { Val = options.WordTableStyle }));
        AppendWordRow(table, sheet.Headers, isHeader: true);
        foreach (var row in sheet.Rows)
            AppendWordRow(table, row, isHeader: false);
        body.Append(table);
        if (options.WordLandscape)
        {
            body.Append(new Wp.SectionProperties(
                new Wp.PageSize { Width = 16838, Height = 11906, Orient = Wp.PageOrientationValues.Landscape }));
        }
        mainPart.Document.Save();
    }

    private sealed record SelectedWorksheet(
        IReadOnlyList<string> Headers,
        IReadOnlyList<IReadOnlyList<string>> Rows,
        IReadOnlyList<IReadOnlyList<object?>> TypedRows);
}
