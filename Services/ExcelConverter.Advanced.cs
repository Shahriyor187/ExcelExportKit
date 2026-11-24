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
    public static Task ToCSVAsync(string inputFile, string outputFile, ExcelConversionOptions options) =>
        ToCSVAsync(inputFile, outputFile, options, CancellationToken.None);

    public static Task ToCSVAsync(string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken) =>
        ConvertWithCancellationAsync(outputFile, cancellationToken,
            path => ToCSVAsyncCore(inputFile, path, options, cancellationToken));

    private static async Task ToCSVAsyncCore(string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(options);
        ValidateCsvDelimiter(options.CsvDelimiter);
        var sheet = ReadWorksheet(inputFile, options, cancellationToken);
        // Keep only one CSV row in the output buffer instead of duplicating the
        // entire workbook as a single StringBuilder in memory.
        await using var writer = new StreamWriter(outputFile, false, Utf8WithoutBom);
        var buffer = new StringBuilder();
        AppendSelectedCsvRow(buffer, sheet.Headers.Select(
            heading => SafeCsvValue(heading, options.CsvProtectFormulas)), options.CsvDelimiter);
        await writer.WriteAsync(buffer.ToString().AsMemory(), cancellationToken).ConfigureAwait(false);
        foreach (var row in sheet.Rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            buffer.Clear();
            AppendSelectedCsvRow(buffer, row.Select(
                value => SafeCsvValue(value, options.CsvProtectFormulas)), options.CsvDelimiter);
            await writer.WriteAsync(buffer.ToString().AsMemory(), cancellationToken).ConfigureAwait(false);
        }
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>Converts the selected worksheet to JSON with optional native numbers and booleans.</summary>
    public static Task ToJsonAsync(string inputFile, string outputFile, ExcelConversionOptions options) =>
        ToJsonAsync(inputFile, outputFile, options, CancellationToken.None);

    public static Task ToJsonAsync(string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken) =>
        ConvertWithCancellationAsync(outputFile, cancellationToken,
            path => ToJsonAsyncCore(inputFile, path, options, cancellationToken));

    private static async Task ToJsonAsyncCore(string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(options);
        var sheet = ReadWorksheet(inputFile, options, cancellationToken);
        // Stream JSON records directly to the file: avoid a second list of
        // dictionaries plus a full serialized JSON string.
        await using var stream = new FileStream(
            outputFile, FileMode.Create, FileAccess.Write, FileShare.None,
            bufferSize: 64 * 1024, useAsync: true);
        using var jsonWriter = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true });
        jsonWriter.WriteStartArray();
        if (options.PreserveJsonTypes)
        {
            foreach (var row in sheet.TypedRows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                jsonWriter.WriteStartObject();
                for (var i = 0; i < sheet.Headers.Count; i++)
                {
                    jsonWriter.WritePropertyName(sheet.Headers[i]);
                    JsonSerializer.Serialize(jsonWriter, row[i], JsonOptions);
                }
                jsonWriter.WriteEndObject();
            }
        }
        else
        {
            foreach (var row in sheet.Rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                jsonWriter.WriteStartObject();
                for (var i = 0; i < sheet.Headers.Count; i++)
                    jsonWriter.WriteString(sheet.Headers[i], row[i]);
                jsonWriter.WriteEndObject();
            }
        }
        jsonWriter.WriteEndArray();
        await jsonWriter.FlushAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>Exports the selected worksheet as a paginated PDF with repeated table headers.</summary>
    public static Task ToPDFAsync(string inputFile, string outputFile, ExcelConversionOptions options) =>
        ToPDFAsync(inputFile, outputFile, options, CancellationToken.None);

    public static Task ToPDFAsync(string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken) =>
        ConvertWithCancellationAsync(outputFile, cancellationToken,
            path => ToPDFAsyncCore(inputFile, path, options, cancellationToken));

    private static Task ToPDFAsyncCore(string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(options);
        if (options.PdfUseWorksheetPrintSettings)
            return WriteWorksheetPrintPdfAsync(inputFile, outputFile, options, cancellationToken);
        if (options.PdfIncludeEmbeddedImages && !options.PdfPreserveCellFormatting)
            throw new ArgumentException(
                "PdfIncludeEmbeddedImages requires PdfPreserveCellFormatting.", nameof(options));
        if (options.PdfPreserveCellFormatting)
            return WriteFormattingAwarePdfAsync(inputFile, outputFile, options, cancellationToken);

        var sheet = ReadWorksheet(inputFile, options, cancellationToken);
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
                            cancellationToken.ThrowIfCancellationRequested();
                            table.Row(row =>
                            {
                                foreach (var value in values)
                                    row.Cell().BorderBottom(0.5, "#dddddd").Padding(5).Text(value);
                            });
                        }
                    });
                });
            }).PublishPdf(outputFile);
            cancellationToken.ThrowIfCancellationRequested();
        }, cancellationToken);
    }

    /// <summary>Converts the selected worksheet to a DOCX table with configurable title and style.</summary>
    public static Task ToWordAsync(string inputFile, string outputFile, ExcelConversionOptions options) =>
        ToWordAsync(inputFile, outputFile, options, CancellationToken.None);

    public static Task ToWordAsync(string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken) =>
        ConvertWithCancellationAsync(outputFile, cancellationToken,
            path => ToWordAsyncCore(inputFile, path, options, cancellationToken));

    private static Task ToWordAsyncCore(string inputFile, string outputFile, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(options);
        var sheet = ReadWorksheet(inputFile, options, cancellationToken);
        return Task.Run(() => WriteStyledWordDocument(sheet, outputFile, options, cancellationToken), cancellationToken);
    }

    // Keep existing files intact on cancellation. Stage only cancellable path exports.
    private static async Task ConvertWithCancellationAsync(string outputFile, CancellationToken cancellationToken,
        Func<string, Task> convert)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled)
        {
            await convert(outputFile).ConfigureAwait(false);
            return;
        }
        ArgumentException.ThrowIfNullOrWhiteSpace(outputFile);
        var fullPath = Path.GetFullPath(outputFile);
        var temporary = Path.Combine(Path.GetDirectoryName(fullPath)!, $".export-{Guid.NewGuid():N}{Path.GetExtension(fullPath)}");
        try
        {
            await convert(temporary).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporary, fullPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary))
                File.Delete(temporary);
        }
    }

    /// <summary>Converts every worksheet to a separate file in the requested format.</summary>
    public static async Task<IReadOnlyList<ExcelConversionResult>> ConvertAllSheetsAsync(
        string inputFile, string outputDirectory, string format, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentException.ThrowIfNullOrWhiteSpace(inputFile);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        ArgumentException.ThrowIfNullOrWhiteSpace(format);
        var normalized = format.ToLowerInvariant();
        if (normalized is not ("csv" or "json" or "pdf" or "docx"))
            throw new ArgumentException("Format must be csv, json, pdf, or docx.", nameof(format));
        if (!File.Exists(inputFile))
            throw new FileNotFoundException("Workbook not found.", inputFile);

        using var workbook = new XLWorkbook(inputFile);
        cancellationToken.ThrowIfCancellationRequested();
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
                    case "csv": await ToCSVAsync(inputFile, destination, options, cancellationToken).ConfigureAwait(false); break;
                    case "json": await ToJsonAsync(inputFile, destination, options, cancellationToken).ConfigureAwait(false); break;
                    case "pdf": await ToPDFAsync(inputFile, destination, options, cancellationToken).ConfigureAwait(false); break;
                    case "docx": await ToWordAsync(inputFile, destination, options, cancellationToken).ConfigureAwait(false); break;
                }
                results.Add(new ExcelConversionResult(inputFile, destination, true, null));
            }
            catch (Exception error) when (error is not OperationCanceledException)
            {
                results.Add(new ExcelConversionResult(inputFile, null, false, $"{sheetName}: {error.Message}"));
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return results;
    }

    /// <summary>Converts several workbooks while reporting individual failures.</summary>
    public static async Task<IReadOnlyList<ExcelConversionResult>> ConvertBatchAsync(
        IEnumerable<string> inputFiles, string outputDirectory, string format, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
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
                    case "csv": await ToCSVAsync(file, path, new ExcelConversionOptions { CsvProtectFormulas = false }, cancellationToken).ConfigureAwait(false); break;
                    case "json": await ToJsonAsync(file, path, new ExcelConversionOptions(), cancellationToken).ConfigureAwait(false); break;
                    case "pdf": await ToPDFAsync(file, path, new ExcelConversionOptions(), cancellationToken).ConfigureAwait(false); break;
                    case "docx": await ToWordAsync(file, path, new ExcelConversionOptions(), cancellationToken).ConfigureAwait(false); break;
                }
                results.Add(new ExcelConversionResult(file, path, true, null));
            }
            catch (Exception error) when (error is not OperationCanceledException)
            {
                results.Add(new ExcelConversionResult(file, null, false, error.Message));
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return results;
    }

    private static SelectedWorksheet ReadWorksheet(string filePath, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath))
            throw new FileNotFoundException("The Excel input file was not found.", filePath);
        ValidateWorkbookInput(filePath, options, cancellationToken);
        using var workbook = new XLWorkbook(filePath);
        cancellationToken.ThrowIfCancellationRequested();
        if (workbook.Worksheets.Count == 0)
            throw new InvalidDataException("The workbook contains no worksheets.");
        if (options.WorksheetIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(options.WorksheetIndex));
        var worksheet = options.WorksheetName is not null
            ? workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, options.WorksheetName, StringComparison.OrdinalIgnoreCase))
            : workbook.Worksheets.Skip(options.WorksheetIndex).FirstOrDefault();
        if (worksheet is null)
            throw new ArgumentException("The selected worksheet does not exist.", nameof(options));
        var usedRange = string.IsNullOrWhiteSpace(options.RangeAddress)
            ? worksheet.RangeUsed()
            : worksheet.Range(options.RangeAddress);
        if (usedRange is null || usedRange.IsEmpty())
            throw new InvalidDataException("The selected worksheet range does not contain any data.");
        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
        var firstColumn = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumn = usedRange.RangeAddress.LastAddress.ColumnNumber;
        ValidateSelectedCells(firstRow, lastRow, firstColumn, lastColumn, options);
        var rawHeaders = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1)
            .Select(col => worksheet.Cell(firstRow, col).GetFormattedString(CultureInfo.InvariantCulture).Trim())
            .ToList();
        var headers = MakeUniqueHeaders(rawHeaders);
        var displayRows = new List<IReadOnlyList<string>>();
        // Only native JSON needs the second, typed representation. CSV, PDF,
        // Word and default JSON should not allocate a duplicate cell matrix.
        List<IReadOnlyList<object?>>? typedRows = options.PreserveJsonTypes
            ? new List<IReadOnlyList<object?>>() : null;

        for (var row = firstRow + 1; row <= lastRow; row++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var display = new List<string>();
            List<object?>? typed = typedRows is null ? null : new List<object?>();
            for (var col = firstColumn; col <= lastColumn; col++)
            {
                var cell = worksheet.Cell(row, col);
                var formatted = cell.GetFormattedString(CultureInfo.InvariantCulture);
                display.Add(formatted);
                if (typed is not null)
                    typed.Add(cell.DataType switch
                    {
                        XLDataType.Number => cell.GetDouble(),
                        XLDataType.Boolean => cell.GetBoolean(),
                        XLDataType.DateTime => cell.GetDateTime(),
                        XLDataType.TimeSpan => cell.GetTimeSpan().ToString(),
                        XLDataType.Blank => null,
                        _ => formatted
                    });
            }
            if (options.SkipEmptyRows && display.All(string.IsNullOrWhiteSpace))
                continue;
            displayRows.Add(display);
            if (typed is not null)
                typedRows!.Add(typed);
        }

        if (headers.Count == 0)
            throw new InvalidDataException("The worksheet has no columns.");
        return new SelectedWorksheet(headers, displayRows,
            typedRows is null
                ? (IReadOnlyList<IReadOnlyList<object?>>)Array.Empty<IReadOnlyList<object?>>()
                : typedRows);
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

    private static void WriteStyledWordDocument(SelectedWorksheet sheet, string outputPath, ExcelConversionOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.WordColumnWidthTwips is <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.WordColumnWidthTwips));

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

        if (options.WordColumnWidthTwips is int width)
        {
            var grid = new Wp.TableGrid();
            foreach (var _ in sheet.Headers)
                grid.Append(new Wp.GridColumn { Width = width.ToString(CultureInfo.InvariantCulture) });
            table.Append(grid);
        }

        AddStyledWordRow(table, sheet.Headers, isHeader: true, options.WordColumnWidthTwips);
        foreach (var row in sheet.Rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            AddStyledWordRow(table, row, isHeader: false, options.WordColumnWidthTwips);
        }
        body.Append(table);

        var section = new Wp.SectionProperties();
        if (!string.IsNullOrWhiteSpace(options.WordHeaderText))
        {
            var header = mainPart.AddNewPart<HeaderPart>();
            header.Header = new Wp.Header(new Wp.Paragraph(new Wp.Run(new Wp.Text(options.WordHeaderText))));
            header.Header.Save();
            section.Append(new Wp.HeaderReference
            {
                Type = Wp.HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(header)
            });
        }

        if (!string.IsNullOrWhiteSpace(options.WordFooterText))
        {
            var footer = mainPart.AddNewPart<FooterPart>();
            footer.Footer = new Wp.Footer(new Wp.Paragraph(new Wp.Run(new Wp.Text(options.WordFooterText))));
            footer.Footer.Save();
            section.Append(new Wp.FooterReference
            {
                Type = Wp.HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(footer)
            });
        }

        if (options.WordLandscape)
        {
            section.Append(new Wp.PageSize
            {
                Width = 16838,
                Height = 11906,
                Orient = Wp.PageOrientationValues.Landscape
            });
        }

        if (section.HasChildren)
            body.Append(section);
        cancellationToken.ThrowIfCancellationRequested();
        mainPart.Document.Save();
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static void AddStyledWordRow(
        Wp.Table table, IReadOnlyList<string> values, bool isHeader, int? columnWidthTwips)
    {
        var tableRow = new Wp.TableRow();
        foreach (var value in values)
        {
            var run = new Wp.Run();
            if (isHeader)
                run.Append(new Wp.RunProperties(new Wp.Bold()));
            run.Append(new Wp.Text(value) { Space = SpaceProcessingModeValues.Preserve });
            var cell = new Wp.TableCell(new Wp.Paragraph(run));
            if (columnWidthTwips is int width)
            {
                cell.PrependChild(new Wp.TableCellProperties(new Wp.TableCellWidth
                {
                    Width = width.ToString(CultureInfo.InvariantCulture),
                    Type = Wp.TableWidthUnitValues.Dxa
                }));
            }
            tableRow.Append(cell);
        }
        table.Append(tableRow);
    }

    private sealed record SelectedWorksheet(
        IReadOnlyList<string> Headers,
        IReadOnlyList<IReadOnlyList<string>> Rows,
        IReadOnlyList<IReadOnlyList<object?>> TypedRows);
}
