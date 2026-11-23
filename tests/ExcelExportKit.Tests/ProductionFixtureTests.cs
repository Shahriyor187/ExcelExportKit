using ClosedXML.Excel;
using ExportFiles.Services;
using System.Text.Json;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class ProductionFixtureTests
{
    private static string Temp(string extension) =>
        Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

    [Fact]
    public async Task InvoiceWithMixedTypesAndQuotesWorksAcrossFormats()
    {
        var input = Temp(".xlsx");
        var csv = Temp(".csv");
        var json = Temp(".json");
        var pdf = Temp(".pdf");
        var word = Temp(".docx");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Invoice");
                sheet.Cell("A1").Value = "Item";
                sheet.Cell("B1").Value = "Quantity";
                sheet.Cell("C1").Value = "Price";
                sheet.Cell("D1").Value = "Paid";
                sheet.Range("A1:D1").Style.Fill.BackgroundColor = XLColor.FromHtml("#244466");
                sheet.Range("A1:D1").Style.Font.FontColor = XLColor.White;
                for (var i = 0; i < 18; i++)
                {
                    var row = i + 2;
                    sheet.Cell(row, 1).Value = i == 0 ? "Books, \"English\"" : "Item " + i;
                    sheet.Cell(row, 2).Value = i + 1;
                    sheet.Cell(row, 3).Value = 4500.25 + i * 100;
                    sheet.Cell(row, 4).Value = i % 2 == 0;
                }
                book.SaveAs(input);
            }
            var options = new ExcelConversionOptions
            {
                WorksheetName = "Invoice",
                PreserveJsonTypes = true,
                PdfPreserveCellFormatting = true,
                MaxWorkbookFileBytes = 10 * 1024 * 1024,
                MaxUncompressedWorkbookBytes = 20 * 1024 * 1024,
                MaxWorkbookArchiveEntries = 150,
                MaxSelectedCells = 100
            };
            await ExcelConverter.ToCSVAsync(input, csv, options);
            await ExcelConverter.ToJsonAsync(input, json, options);
            await ExcelConverter.ToPDFAsync(input, pdf, options);
            await ExcelConverter.ToWordAsync(input, word, options);

            Assert.Contains("\"Books, \"\"English\"\"\"", await File.ReadAllTextAsync(csv));
            using var jsonDoc = JsonDocument.Parse(await File.ReadAllTextAsync(json));
            Assert.Equal(18, jsonDoc.RootElement.GetArrayLength());
            Assert.True(jsonDoc.RootElement[0].GetProperty("Paid").GetBoolean());
            Assert.Equal(4500.25, jsonDoc.RootElement[0].GetProperty("Price").GetDouble());
            Assert.True(new FileInfo(pdf).Length > 100);
            Assert.True(new FileInfo(word).Length > 100);
        }
        finally
        {
            foreach (var path in new[] { input, csv, json, pdf, word })
                if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task MultiThousandRowReportPreservesRecordCount()
    {
        var input = Temp(".xlsx");
        var json = Temp(".json");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Report");
                sheet.Cell("A1").Value = "Id";
                sheet.Cell("B1").Value = "Student";
                sheet.Cell("C1").Value = "Score";
                for (var i = 1; i <= 2_000; i++)
                {
                    sheet.Cell(i + 1, 1).Value = i;
                    sheet.Cell(i + 1, 2).Value = "Student " + i;
                    sheet.Cell(i + 1, 3).Value = 900 + i % 700;
                }
                book.SaveAs(input);
            }
            await ExcelConverter.ToJsonAsync(input, json, new ExcelConversionOptions
            {
                WorksheetName = "Report",
                PreserveJsonTypes = true,
                MaxSelectedCells = 6_003
            });
            using var parsed = JsonDocument.Parse(await File.ReadAllTextAsync(json));
            Assert.Equal(2_000, parsed.RootElement.GetArrayLength());
            Assert.Equal(2_000, parsed.RootElement[1_999].GetProperty("Id").GetInt32());
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(json)) File.Delete(json);
        }
    }
}
