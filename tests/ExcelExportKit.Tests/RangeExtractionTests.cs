using ClosedXML.Excel;
using ExportFiles.Services;
using System.Text.Json;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class RangeExtractionTests
{
    [Fact]
    public async Task ConvertsOnlyRequestedRangeAndSkipsBlankRows()
    {
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Results");
                sheet.Cell("A1").Value = "Ignored";
                sheet.Cell("B3").Value = "Student";
                sheet.Cell("C3").Value = "Score";
                sheet.Cell("B4").Value = "Amina";
                sheet.Cell("C4").Value = 1500;
                sheet.Cell("B6").Value = "Ali";
                sheet.Cell("C6").Value = 1410;
                book.SaveAs(input);
            }

            await ExcelConverter.ToJsonAsync(input, output, new ExcelConversionOptions
            {
                WorksheetName = "Results",
                RangeAddress = "B3:C6",
                SkipEmptyRows = true,
                PreserveJsonTypes = true
            });

            using var json = JsonDocument.Parse(await File.ReadAllTextAsync(output));
            Assert.Equal(2, json.RootElement.GetArrayLength());
            Assert.Equal("Amina", json.RootElement[0].GetProperty("Student").GetString());
            Assert.Equal(1500, json.RootElement[0].GetProperty("Score").GetDouble());
            Assert.Equal("Ali", json.RootElement[1].GetProperty("Student").GetString());
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public async Task SupportsSelectedRangeInPdfMode()
    {
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Results");
                sheet.Cell("A1").Value = "Ignored";
                sheet.Cell("B2").Value = "Student";
                sheet.Cell("B3").Value = "Amina";
                book.SaveAs(input);
            }
            await ExcelConverter.ToPDFAsync(input, output, new ExcelConversionOptions
            {
                RangeAddress = "B2:B3",
                PdfPreserveCellFormatting = true
            });
            Assert.True(new FileInfo(output).Length > 100);
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }
}
