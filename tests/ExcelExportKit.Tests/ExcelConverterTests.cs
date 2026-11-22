using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using ExportFiles.Services;
using System.Text.Json;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class ExcelConverterTests
{
    [Fact]
    public async Task ToJsonAsync_UsesHeaderCellsAsProperties()
    {
        var input = CreateWorkbook();
        var output = CreateTempFile(".json");

        try
        {
            await ExcelConverter.ToJsonAsync(input, output);

            using var json = JsonDocument.Parse(await File.ReadAllTextAsync(output));
            var firstRecord = json.RootElement[0];

            Assert.Equal("Amina", firstRecord.GetProperty("Name").GetString());
            Assert.Equal("Uzbekistan", firstRecord.GetProperty("Country").GetString());
            Assert.Equal("Ali", json.RootElement[1].GetProperty("Name").GetString());
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

    [Fact]
    public async Task ToCSVAsync_EscapesCsvFields()
    {
        var input = CreateWorkbook(includeQuotedField: true);
        var output = CreateTempFile(".csv");

        try
        {
            await ExcelConverter.ToCSVAsync(input, output);

            var csv = await File.ReadAllTextAsync(output);
            Assert.Contains("Name,Country,Note", csv);
            Assert.Contains("\"Amina, M.\",Uzbekistan,\"said \"\"hello\"\"\"", csv);
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

    [Fact]
    public async Task ToPDFAsync_CreatesPdfFile()
    {
        var input = CreateWorkbook();
        var output = CreateTempFile(".pdf");

        try
        {
            await ExcelConverter.ToPDFAsync(input, output);

            var bytes = await File.ReadAllBytesAsync(output);
            Assert.True(bytes.Length > 100);
            Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(bytes, 0, 5));
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

    [Fact]
    public async Task ToWordAsync_CreatesOneTableWithHeaderAndDataRows()
    {
        var input = CreateWorkbook();
        var output = CreateTempFile(".docx");

        try
        {
            await ExcelConverter.ToWordAsync(input, output);

            using var document = WordprocessingDocument.Open(output, false);
            var table = Assert.Single(document.MainDocumentPart!.Document!.Body!.Elements<
                DocumentFormat.OpenXml.Wordprocessing.Table>());

            Assert.Equal(3, table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Count());
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

    [Fact]
    public async Task ToJsonAsync_RejectsMissingInputFile()
    {
        var output = CreateTempFile(".json");

        try
        {
            await Assert.ThrowsAsync<FileNotFoundException>(
                () => ExcelConverter.ToJsonAsync("missing-workbook.xlsx", output));
        }
        finally
        {
            DeleteIfExists(output);
        }
    }


    [Fact]
    public async Task ToJsonAsync_MakesDuplicateHeadersUnique()
    {
        var input = CreateTempFile(".xlsx");
        var output = CreateTempFile(".json");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell(1, 1).Value = "Name";
                sheet.Cell(1, 2).Value = "Name";
                sheet.Cell(2, 1).Value = "Amina";
                sheet.Cell(2, 2).Value = "Ali";
                workbook.SaveAs(input);
            }

            await ExcelConverter.ToJsonAsync(input, output);
            using var json = JsonDocument.Parse(File.ReadAllText(output));
            var record = json.RootElement[0];

            Assert.Equal("Amina", record.GetProperty("Name").GetString());
            Assert.Equal("Ali", record.GetProperty("Name_2").GetString());
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

    [Fact]
    public async Task ToJsonAsync_RejectsEmptyWorksheet()
    {
        var input = CreateTempFile(".xlsx");
        var output = CreateTempFile(".json");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Empty");
                workbook.SaveAs(input);
            }

            await Assert.ThrowsAsync<InvalidDataException>(
                () => ExcelConverter.ToJsonAsync(input, output));
        }
        finally
        {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }


    [Fact]
    public async Task ToJsonAsync_SelectsNamedWorksheetAndPreservesNativeTypes()
    {
        var input = CreateTempFile(".xlsx");
        var output = CreateTempFile(".json");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Other").Cell("A1").Value = "NotSelected";
                var sheet = workbook.AddWorksheet("Results");
                sheet.Cell("A1").Value = "Score";
                sheet.Cell("B1").Value = "Passed";
                sheet.Cell("A2").Value = 1450;
                sheet.Cell("B2").Value = true;
                workbook.SaveAs(input);
            }

            await ExcelConverter.ToJsonAsync(input, output, new ExcelConversionOptions
            {
                WorksheetName = "Results",
                PreserveJsonTypes = true
            });
            using var json = JsonDocument.Parse(File.ReadAllText(output));
            Assert.Equal(1450, json.RootElement[0].GetProperty("Score").GetDouble());
            Assert.True(json.RootElement[0].GetProperty("Passed").GetBoolean());
        }
        finally { DeleteIfExists(input); DeleteIfExists(output); }
    }

    [Fact]
    public async Task ToCsvAsync_SelectsWorksheetByIndexAndCustomDelimiter()
    {
        var input = CreateTempFile(".xlsx");
        var output = CreateTempFile(".csv");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Ignore").Cell("A1").Value = "Ignore";
                var sheet = workbook.AddWorksheet("Second");
                sheet.Cell("A1").Value = "Name";
                sheet.Cell("B1").Value = "Score";
                sheet.Cell("A2").Value = "Ali";
                sheet.Cell("B2").Value = 1500;
                workbook.SaveAs(input);
            }
            await ExcelConverter.ToCSVAsync(input, output, new ExcelConversionOptions
            {
                WorksheetIndex = 1,
                CsvDelimiter = ';'
            });
            Assert.Contains("Name;Score", File.ReadAllText(output));
            Assert.Contains("Ali;1500", File.ReadAllText(output));
        }
        finally { DeleteIfExists(input); DeleteIfExists(output); }
    }

    [Fact]
    public async Task ToPdfAsync_UsesConfiguredTitleAndLandscapeLayout()
    {
        var input = CreateWorkbook();
        var output = CreateTempFile(".pdf");
        try
        {
            await ExcelConverter.ToPDFAsync(input, output, new ExcelConversionOptions
            {
                DocumentTitle = "Quarterly results",
                PdfLandscape = true
            });
            var bytes = File.ReadAllBytes(output);
            Assert.True(bytes.Length > 100);
            Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(bytes, 0, 5));
        }
        finally { DeleteIfExists(input); DeleteIfExists(output); }
    }

    [Fact]
    public async Task ToWordAsync_RespectsCustomDocumentTitle()
    {
        var input = CreateWorkbook();
        var output = CreateTempFile(".docx");
        try
        {
            await ExcelConverter.ToWordAsync(input, output,
                new ExcelConversionOptions { DocumentTitle = "Report" });
            using var document = WordprocessingDocument.Open(output, false);
            var text = document.MainDocumentPart!.Document!.Body!.InnerText;
            Assert.Contains("Report", text);
            Assert.Contains("Amina", text);
        }
        finally { DeleteIfExists(input); DeleteIfExists(output); }
    }

    [Fact]
    public async Task ConvertAllSheetsAsync_WritesEachSheetToItsOwnFile()
    {
        var input = CreateTempFile(".xlsx");
        var folder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var first = workbook.AddWorksheet("January");
                first.Cell("A1").Value = "Name";
                first.Cell("A2").Value = "Ali";
                var second = workbook.AddWorksheet("February");
                second.Cell("A1").Value = "Name";
                second.Cell("A2").Value = "Amina";
                workbook.SaveAs(input);
            }
            var results = await ExcelConverter.ConvertAllSheetsAsync(input, folder, "json");
            Assert.Equal(2, results.Count);
            Assert.All(results, result => Assert.True(result.Success, result.Error));
            Assert.Equal(2, Directory.GetFiles(folder, "*.json").Length);
        }
        finally
        {
            DeleteIfExists(input);
            if (Directory.Exists(folder)) Directory.Delete(folder, true);
        }
    }

    [Fact]
    public async Task ConvertBatchAsync_ReportsIndividualFileErrors()
    {
        var good = CreateWorkbook();
        var missing = CreateTempFile(".xlsx");
        var folder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        try
        {
            var results = await ExcelConverter.ConvertBatchAsync([good, missing], folder, "csv");
            Assert.Equal(2, results.Count);
            Assert.True(results[0].Success);
            Assert.False(results[1].Success);
            Assert.NotNull(results[1].Error);
        }
        finally
        {
            DeleteIfExists(good);
            if (Directory.Exists(folder)) Directory.Delete(folder, true);
        }
    }

    private static string CreateWorkbook(bool includeQuotedField = false)
    {
        var path = CreateTempFile(".xlsx");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");

        worksheet.Cell(1, 1).Value = "Name";
        worksheet.Cell(1, 2).Value = "Country";

        if (includeQuotedField)
        {
            worksheet.Cell(1, 3).Value = "Note";
        }

        worksheet.Cell(2, 1).Value = includeQuotedField ? "Amina, M." : "Amina";
        worksheet.Cell(2, 2).Value = "Uzbekistan";

        if (includeQuotedField)
        {
            worksheet.Cell(2, 3).Value = "said \"hello\"";
        }

        worksheet.Cell(3, 1).Value = "Ali";
        worksheet.Cell(3, 2).Value = "Uzbekistan";

        workbook.SaveAs(path);
        return path;
    }


    [Fact]
    public async Task ToCsvAsync_ProtectsFormulaTextWhenRequested()
    {
        var input = CreateTempFile(".xlsx");
        var output = CreateTempFile(".csv");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Data");
                worksheet.Cell("A1").Value = "Formula";
                worksheet.Cell("A2").Value = "=1+2";
                workbook.SaveAs(input);
            }
            await ExcelConverter.ToCSVAsync(input, output,
                new ExcelConversionOptions { CsvProtectFormulas = true });
            Assert.Contains("'=1+2", File.ReadAllText(output));
        }
        finally { DeleteIfExists(input); DeleteIfExists(output); }
    }

    [Fact]
    public async Task ToWordAsync_WritesLandscapePageSettings()
    {
        var input = CreateWorkbook();
        var output = CreateTempFile(".docx");
        try
        {
            await ExcelConverter.ToWordAsync(input, output,
                new ExcelConversionOptions { WordLandscape = true });
            using var document = WordprocessingDocument.Open(output, false);
            var section = Assert.Single(document.MainDocumentPart!.Document!.Body!
                .Elements<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>());
            Assert.NotNull(section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>());
        }
        finally { DeleteIfExists(input); DeleteIfExists(output); }
    }

    private static string CreateTempFile(string extension) =>
        Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}{extension}");

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
