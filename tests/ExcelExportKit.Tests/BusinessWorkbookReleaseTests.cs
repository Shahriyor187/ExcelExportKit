using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using ExportFiles.Services;
using System.Text.Json;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class BusinessWorkbookReleaseTests
{
    private static string FilePath(string extension) =>
        Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

    [Fact]
    public async Task MultiSheetInvoiceWithUnicodeAndTotalsExportsSelectedSheetAcrossFourFormats()
    {
        var source = FilePath(".xlsx");
        var csv = FilePath(".csv");
        var json = FilePath(".json");
        var pdf = FilePath(".pdf");
        var docx = FilePath(".docx");
        try
        {
            using (var book = new XLWorkbook())
            {
                var summary = book.AddWorksheet("Summary");
                summary.Cell("A1").Value = "Other worksheet";
                summary.Cell("A2").Value = "Do not export";

                var invoice = book.AddWorksheet("Invoice 2026");
                foreach (var (column, title) in new[]
                {
                    (1, "Item"), (2, "Quantity"), (3, "Unit Price"),
                    (4, "Paid"), (5, "Notes")
                })
                    invoice.Cell(1, column).Value = title;
                invoice.Range("A1:E1").Style.Fill.BackgroundColor = XLColor.FromHtml("#25466A");
                invoice.Range("A1:E1").Style.Font.FontColor = XLColor.White;
                invoice.Range("A1:E1").Style.Font.Bold = true;

                for (var i = 0; i < 48; i++)
                {
                    var row = i + 2;
                    invoice.Cell(row, 1).Value = i == 0 ? "Toshkent, \"Service\"" :
                        "Mahsulot " + (i + 1);
                    invoice.Cell(row, 2).Value = i + 1;
                    invoice.Cell(row, 3).Value = 12500.75 + i * 125;
                    invoice.Cell(row, 4).Value = i % 2 == 0;
                    invoice.Cell(row, 5).Value = i == 0 ? "To'lov \n tasdiqlandi" : "Standard";
                }
                invoice.Column(1).Width = 31;
                invoice.Column(2).Width = 16;
                invoice.Column(3).Width = 19;
                invoice.Column(4).Width = 14;
                invoice.Column(5).Width = 24;
                book.SaveAs(source);
            }

            var options = new ExcelConversionOptions
            {
                WorksheetName = "Invoice 2026",
                PreserveJsonTypes = true,
                PdfPreserveCellFormatting = true,
                SkipEmptyRows = true,
                MaxWorkbookFileBytes = 10 * 1024 * 1024,
                MaxUncompressedWorkbookBytes = 20 * 1024 * 1024,
                MaxWorkbookArchiveEntries = 512,
                MaxSelectedCells = 1_000
            };
            await ExcelConverter.ToCSVAsync(source, csv, options);
            await ExcelConverter.ToJsonAsync(source, json, options);
            await ExcelConverter.ToPDFAsync(source, pdf, options);
            await ExcelConverter.ToWordAsync(source, docx, options);

            var csvText = await File.ReadAllTextAsync(csv);
            Assert.Contains("\"Toshkent, \"\"Service\"\"\"", csvText);
            Assert.DoesNotContain("Other worksheet", csvText);
            using var document = JsonDocument.Parse(await File.ReadAllTextAsync(json));
            Assert.Equal(48, document.RootElement.GetArrayLength());
            Assert.Equal("Toshkent, \"Service\"",
                document.RootElement[0].GetProperty("Item").GetString());
            Assert.Equal(48, document.RootElement[47].GetProperty("Quantity").GetInt32());
            Assert.Equal(12500.75, document.RootElement[0].GetProperty("Unit Price").GetDouble(), 5);
            Assert.True(document.RootElement[0].GetProperty("Paid").GetBoolean());
            Assert.Contains("tasdiqlandi", document.RootElement[0].GetProperty("Notes").GetString());

            using (var file = File.OpenRead(pdf))
            {
                var signature = new byte[5];
                Assert.Equal(5, file.Read(signature));
                Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(signature));
            }
            using (var word = WordprocessingDocument.Open(docx, false))
            {
                var text = word.MainDocumentPart!.Document!.Body!.InnerText;
                Assert.Contains("Mahsulot 48", text);
                Assert.DoesNotContain("Do not export", text);
            }
        }
        finally
        {
            foreach (var path in new[] { source, csv, json, pdf, docx })
                if (File.Exists(path)) File.Delete(path);
        }
    }
}
