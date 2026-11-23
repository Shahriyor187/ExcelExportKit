using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ExportFiles.Services;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class PdfEmbeddedImageTests
{
    // A valid 1x1 pixel PNG fixture; the test fixture does not depend on local image files.
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/x4sAAAAASUVORK5CYII=");

    [Fact]
    public async Task CanEmbedPngAnchoredToAnUnmergedCell()
    {
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        var output = Path.ChangeExtension(input, ".pdf");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Report");
                sheet.Cell("A1").Value = "Student";
                sheet.Cell("B1").Value = "Photo";
                sheet.Cell("A2").Value = "Amina";
                using var image = new MemoryStream(Png);
                sheet.AddPicture(image, XLPictureFormat.Png)
                    .MoveTo(sheet.Cell("B2"));
                book.SaveAs(input);
            }

            var options = new ExcelConversionOptions
            {
                PdfPreserveCellFormatting = true,
                PdfIncludeEmbeddedImages = true
            };
            var snapshot = ExcelConverter.ReadPdfFormatting(input, options);
            Assert.Equal(Png, snapshot.Rows[0].Cells[1].Picture!.Bytes);
            await ExcelConverter.ToPDFAsync(input, output, options);
            Assert.True(new FileInfo(output).Length > 100);
            using var stream = File.OpenRead(output);
            var signature = new byte[5];
            Assert.Equal(5, stream.Read(signature));
            Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(signature));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void FailsExplicitlyWhenPictureIsOutsideSelectedRange()
    {
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Report");
                sheet.Cell("A1").Value = "Header";
                sheet.Cell("A2").Value = "Value";
                sheet.Cell("B2").Value = "Image";
                using var picture = new MemoryStream(Png);
                sheet.AddPicture(picture, XLPictureFormat.Png).MoveTo(sheet.Cell("C2"));
                book.SaveAs(input);
            }
            Assert.Throws<NotSupportedException>(() =>
                ExcelConverter.ReadPdfFormatting(input, new ExcelConversionOptions
                {
                    PdfPreserveCellFormatting = true,
                    PdfIncludeEmbeddedImages = true,
                    RangeAddress = "A1:B2"
                }));
        }
        finally { if (File.Exists(input)) File.Delete(input); }
    }

    [Fact]
    public void RejectsInvalidImageLimits()
    {
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Data");
                sheet.Cell("A1").Value = "Name";
                sheet.Cell("A2").Value = "Test";
                book.SaveAs(input);
            }
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                ExcelConverter.ReadPdfFormatting(input, new ExcelConversionOptions
                {
                    PdfIncludeEmbeddedImages = true,
                    PdfPreserveCellFormatting = true,
                    PdfMaxEmbeddedImageBytes = 20
                }));
        }
        finally { if (File.Exists(input)) File.Delete(input); }
    }
}
