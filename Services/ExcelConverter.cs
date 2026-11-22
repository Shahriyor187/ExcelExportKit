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

/// <summary>
/// Converts the first worksheet of an Excel workbook to CSV, JSON, PDF, or Word.
/// </summary>
/// <remarks>
/// The first used row is treated as the header row. Blank or duplicate headers
/// are replaced with unique column names. Cell values are exported using their
/// formatted display text.
/// </remarks>
public static class ExcelConverter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    /// <summary>
    /// Converts the first worksheet of an Excel workbook to CSV.
    /// Fields containing commas, quotes, or line breaks are escaped.
    /// </summary>
    public static async Task ToCSVAsync(string inputFile, string outputFile)
    {
        var sheet = ReadFirstWorksheet(inputFile);
        var builder = new StringBuilder();

        AppendCsvRow(builder, sheet.Headers);

        foreach (var row in sheet.Rows)
        {
            AppendCsvRow(builder, row);
        }

        await File.WriteAllTextAsync(outputFile, builder.ToString(), Utf8WithoutBom)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Converts the first worksheet of an Excel workbook to a JSON array of objects.
    /// The first used row supplies the property names.
    /// </summary>
    public static async Task ToJsonAsync(string inputFile, string outputFile)
    {
        var sheet = ReadFirstWorksheet(inputFile);
        var records = new List<Dictionary<string, string>>(sheet.Rows.Count);

        foreach (var row in sheet.Rows)
        {
            var record = new Dictionary<string, string>(StringComparer.Ordinal);

            for (var columnIndex = 0; columnIndex < sheet.Headers.Count; columnIndex++)
            {
                record[sheet.Headers[columnIndex]] = row[columnIndex];
            }

            records.Add(record);
        }

        var json = JsonSerializer.Serialize(records, JsonOptions);
        await File.WriteAllTextAsync(outputFile, json, Utf8WithoutBom)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Converts the first worksheet of an Excel workbook to a paginated PDF document.
    /// </summary>
    public static Task ToPDFAsync(string inputFile, string outputFile)
    {
        var sheet = ReadFirstWorksheet(inputFile);

        return Task.Run(() =>
        {
            TerraPDF.Core.Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSize.A4);
                    page.Margin(2, Unit.Centimetre);
                    page.DefaultTextStyle(style => style.FontSize(9));

                    page.Header()
                        .Text("Excel worksheet export")
                        .Bold()
                        .FontSize(14);

                    page.Content().Column(column =>
                    {
                        column.Spacing(4);
                        column.Item()
                            .Text(string.Join("  |  ", sheet.Headers))
                            .Bold();

                        foreach (var row in sheet.Rows)
                        {
                            column.Item().Text(string.Join("  |  ", row));
                        }
                    });
                });
            }).PublishPdf(outputFile);
        });
    }

    /// <summary>
    /// Converts the first worksheet of an Excel workbook to a Word document
    /// containing a single table with a header row.
    /// </summary>
    /// <remarks>
    /// This method preserves the original API name. Prefer <see cref="ToWordAsync"/>
    /// in new code.
    /// </remarks>
    public static Task ToWord(string inputFile, string outputFile) =>
        ToWordAsync(inputFile, outputFile);

    /// <summary>
    /// Converts the first worksheet of an Excel workbook to a Word document
    /// containing a single table with a header row.
    /// </summary>
    public static Task ToWordAsync(string inputFile, string outputFile)
    {
        var sheet = ReadFirstWorksheet(inputFile);

        return Task.Run(() => WriteWordDocument(sheet, outputFile));
    }

    private static WorksheetData ReadFirstWorksheet(string inputFile)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputFile);

        if (!File.Exists(inputFile))
        {
            throw new FileNotFoundException("The Excel input file was not found.", inputFile);
        }

        using var workbook = new XLWorkbook(inputFile);
        var worksheet = workbook.Worksheets.FirstOrDefault()
            ?? throw new InvalidDataException("The Excel workbook does not contain any worksheets.");

        var usedRange = worksheet.RangeUsed()
            ?? throw new InvalidDataException("The first worksheet does not contain any data.");

        var firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        var lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
        var firstColumn = usedRange.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumn = usedRange.RangeAddress.LastAddress.ColumnNumber;
        var columnCount = lastColumn - firstColumn + 1;

        var rawHeaders = new List<string>(columnCount);
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            rawHeaders.Add(worksheet.Cell(firstRow, column)
                .GetFormattedString(CultureInfo.InvariantCulture).Trim());
        }

        var headers = MakeUniqueHeaders(rawHeaders);
        var rows = new List<IReadOnlyList<string>>(Math.Max(0, lastRow - firstRow));

        for (var row = firstRow + 1; row <= lastRow; row++)
        {
            var values = new List<string>(columnCount);

            for (var column = firstColumn; column <= lastColumn; column++)
            {
                values.Add(worksheet.Cell(row, column)
                    .GetFormattedString(CultureInfo.InvariantCulture));
            }

            rows.Add(values);
        }

        return new WorksheetData(headers, rows);
    }

    private static IReadOnlyList<string> MakeUniqueHeaders(IReadOnlyList<string> rawHeaders)
    {
        var headers = new List<string>(rawHeaders.Count);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var index = 0; index < rawHeaders.Count; index++)
        {
            var baseName = string.IsNullOrWhiteSpace(rawHeaders[index])
                ? $"Column{index + 1}"
                : rawHeaders[index];

            var candidate = baseName;
            var suffix = 2;

            while (!used.Add(candidate))
            {
                candidate = $"{baseName}_{suffix++}";
            }

            headers.Add(candidate);
        }

        return headers;
    }

    private static void AppendCsvRow(StringBuilder builder, IEnumerable<string> values)
    {
        builder.AppendJoin(',', values.Select(EscapeCsvValue));
        builder.AppendLine();
    }

    private static string EscapeCsvValue(string value)
    {
        if (value.IndexOfAny([',', '"', '\r', '\n']) < 0)
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";
    }

    private static void WriteWordDocument(WorksheetData sheet, string outputFile)
    {
        using var wordDocument = WordprocessingDocument.Create(
            outputFile,
            WordprocessingDocumentType.Document);

        var mainPart = wordDocument.AddMainDocumentPart();
        var document = new Wp.Document(new Wp.Body());
        mainPart.Document = document;

        var table = new Wp.Table();
        table.AppendChild(new Wp.TableProperties(
            new Wp.TableBorders(
                new Wp.TopBorder { Val = Wp.BorderValues.Single, Size = 4 },
                new Wp.BottomBorder { Val = Wp.BorderValues.Single, Size = 4 },
                new Wp.LeftBorder { Val = Wp.BorderValues.Single, Size = 4 },
                new Wp.RightBorder { Val = Wp.BorderValues.Single, Size = 4 },
                new Wp.InsideHorizontalBorder { Val = Wp.BorderValues.Single, Size = 4 },
                new Wp.InsideVerticalBorder { Val = Wp.BorderValues.Single, Size = 4 })));

        AppendWordRow(table, sheet.Headers, isHeader: true);

        foreach (var row in sheet.Rows)
        {
            AppendWordRow(table, row, isHeader: false);
        }

        document.Body!.Append(table);
        document.Save();
    }

    private static void AppendWordRow(
        Wp.Table table,
        IReadOnlyList<string> values,
        bool isHeader)
    {
        var tableRow = new Wp.TableRow();

        foreach (var value in values)
        {
            var run = new Wp.Run();

            if (isHeader)
            {
                run.Append(new Wp.RunProperties(new Wp.Bold()));
            }

            run.Append(new Wp.Text(value)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

            tableRow.Append(new Wp.TableCell(new Wp.Paragraph(run)));
        }

        table.Append(tableRow);
    }

    private static readonly Encoding Utf8WithoutBom =
        new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    private sealed record WorksheetData(
        IReadOnlyList<string> Headers,
        IReadOnlyList<IReadOnlyList<string>> Rows);
}
