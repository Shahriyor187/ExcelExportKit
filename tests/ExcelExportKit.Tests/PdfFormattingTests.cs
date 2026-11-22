using ClosedXML.Excel;
using ExportFiles.Services;
using System.Text;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class PdfFormattingTests
{
    [Fact]
    public async Task StyledPdfSupportsWorksheetSelectionAndCellFormatting()
    {
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Ignore").Cell("A1").Value = "Not this";
                var sheet = workbook.AddWorksheet("Results");
                sheet.Column(1).Width = 26;
                sheet.Column(2).Width = 10;
                sheet.Cell("A1").Value = "Student";
                sheet.Cell("B1").Value = "SAT Score";
                sheet.Range("A1:B1").Style.Font.Bold = true;
                sheet.Range("A1:B1").Style.Fill.BackgroundColor = XLColor.FromHtml("#1A4A8A");
                sheet.Range("A1:B1").Style.Font.FontColor = XLColor.White;
                sheet.Cell("A2").Value = "Amina";
                sheet.Cell("B2").Value = 1500;
                sheet.Cell("B2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                sheet.Cell("B2").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF5E9");
                workbook.SaveAs(input);
            }

            await ExcelConverter.ToPDFAsync(input, output, new ExcelConversionOptions
            {
                WorksheetName = "Results",
                PdfPreserveCellFormatting = true,
                PdfLandscape = true,
                DocumentTitle = "SAT Results"
            });

            var bytes = await File.ReadAllBytesAsync(output);
            Assert.True(bytes.Length > 100);
            Assert.Equal("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public async Task StyledPdfRejectsMissingWorksheet()
    {
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Data").Cell("A1").Value = "Name";
                workbook.SaveAs(input);
            }

            await Assert.ThrowsAsync<ArgumentException>(() =>
                ExcelConverter.ToPDFAsync(input, output, new ExcelConversionOptions
                {
                    WorksheetName = "Missing",
                    PdfPreserveCellFormatting = true
                }));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }
}
