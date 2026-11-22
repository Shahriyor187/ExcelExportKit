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
            var table = Assert.Single(document.MainDocumentPart!.Document.Body!.Elements<
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
