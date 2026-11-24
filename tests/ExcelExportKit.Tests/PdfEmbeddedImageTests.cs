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
    public void SkipEmptyRowsDoesNotDiscardImageOnlyRow()
    {
        var input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Report");
                sheet.Cell("A1").Value = "Name";
                sheet.Cell("B1").Value = "Photo";
                sheet.Cell("A3").Value = "Second student";
                using var picture = new MemoryStream(Png);
                sheet.AddPicture(picture, XLPictureFormat.Png)
                    .MoveTo(sheet.Cell("B2"));
                book.SaveAs(input);
            }
            var snapshot = ExcelConverter.ReadPdfFormatting(input, new ExcelConversionOptions
            {
                PdfPreserveCellFormatting = true,
                PdfIncludeEmbeddedImages = true,
                SkipEmptyRows = true,
                RangeAddress = "A1:B3"
            });
            Assert.Equal(2, snapshot.Rows.Count);
            Assert.NotNull(snapshot.Rows[0].Cells[1].Picture);
        }
        finally { if (File.Exists(input)) File.Delete(input); }
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

    // Valid, tiny compressed PNGs with real pixel dimensions; no huge
    // allocations or external image dependencies are needed for this test.
    [Theory]
    [InlineData("iVBORw0KGgoAAAANSUhEUgAAEAEAAAACCAIAAAC4FeW7AAAAL0lEQVR42u3BMQEAAADCoPVPbQlPoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgJcBYAgAAUIdNJYAAAAASUVORK5CYII=", 4097, 2)]
    [InlineData("iVBORw0KGgoAAAANSUhEUgAAAAIAABABCAIAAAAcnnQTAAAAM0lEQVR42u3BgQAAAADDoPlTX+EAVQEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACnAXAHAAE9ArVYAAAAAElFTkSuQmCC", 2, 4097)]
    public async Task ResizingAnOversizedOriginalImageDoesNotBypassLimits(string png, int width, int height)
    {
        using var input = new MemoryStream();
        using (var book = new XLWorkbook())
        {
            var sheet = book.AddWorksheet("Data");
            sheet.Cell("A1").Value = "Name";
            sheet.Cell("B1").Value = "Photo";
            sheet.Cell("A2").Value = "Example";
            using var image = new MemoryStream(Convert.FromBase64String(png));
            var picture = sheet.AddPicture(image, XLPictureFormat.Png).MoveTo(sheet.Cell("B2"));
            Assert.Equal(width, picture.OriginalWidth);
            Assert.Equal(height, picture.OriginalHeight);
            picture.Width = 64;
            picture.Height = 32;
            book.SaveAs(input);
        }
        input.Position = 0;
        using (var loaded = new XLWorkbook(input))
        {
            var picture = Assert.Single(loaded.Worksheet(1).Pictures);
            Assert.Equal(width, picture.OriginalWidth);
            Assert.Equal(height, picture.OriginalHeight);
            Assert.Equal(64, picture.Width);
            Assert.Equal(32, picture.Height);
        }
        input.Position = 0;
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            ExcelConverter.ToPDFAsync(input, output, new ExcelConversionOptions
            {
                RangeAddress = "A1:B2",
                PdfPreserveCellFormatting = true,
                PdfIncludeEmbeddedImages = true
            }));
        Assert.Contains("original or displayed dimensions", error.Message);
        Assert.Equal(0, output.Length);
        Assert.True(input.CanRead);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task OriginalImageAtDimensionLimitIsAccepted()
    {
        using var input = new MemoryStream();
        using (var book = new XLWorkbook())
        {
            var sheet = book.AddWorksheet("Data");
            sheet.Cell("A1").Value = "Name";
            sheet.Cell("B1").Value = "Photo";
            sheet.Cell("A2").Value = "Example";
            using var image = new MemoryStream(Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAEAAAAAACCAIAAABX146FAAAAL0lEQVR42u3BMQEAAADCoPVPbQhfoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgN8AYAIAAbpK7usAAAAASUVORK5CYII="));
            var picture = sheet.AddPicture(image, XLPictureFormat.Png).MoveTo(sheet.Cell("B2"));
            picture.Width = 64;
            picture.Height = 32;
            book.SaveAs(input);
        }
        input.Position = 0;
        using var output = new MemoryStream();
        await ExcelConverter.ToPDFAsync(input, output, new ExcelConversionOptions
        {
            RangeAddress = "A1:B2",
            PdfPreserveCellFormatting = true,
            PdfIncludeEmbeddedImages = true
        });
        Assert.True(output.Length > 100);
        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(output.ToArray(), 0, 5));
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
