using ClosedXML.Excel;
using ExportFiles.Services;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class PdfPrintLayoutTests
{
    private static string Temp(string extension) =>
        Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

    [Fact]
    public async Task RespectsPrintAreaOrientationMarginsScaleAndBothManualBreakDirections()
    {
        var input = Temp(".xlsx");
        var output = Temp(".pdf");
        try
        {
            using (var book = new XLWorkbook())
            {
                var ws = book.AddWorksheet("Print");
                ws.Cell("A1").Value = "Student";
                ws.Cell("B1").Value = "Score";
                for (var row = 2; row <= 8; row++)
                {
                    ws.Cell(row, 1).Value = "Student " + row;
                    ws.Cell(row, 2).Value = 1200 + row;
                }
                ws.Cell("C1").Value = "Outside";
                ws.PageSetup.PrintAreas.Add("A1:B8");
                ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                ws.PageSetup.Margins.Left = 0.35;
                ws.PageSetup.Margins.Right = 0.45;
                ws.PageSetup.Margins.Top = 0.6;
                ws.PageSetup.Margins.Bottom = 0.7;
                ws.PageSetup.Scale = 85;
                ws.PageSetup.AddHorizontalPageBreak(4);
                ws.PageSetup.AddVerticalPageBreak(1);
                book.SaveAs(input);
            }

            var options = new ExcelConversionOptions
            {
                WorksheetName = "Print",
                PdfPreserveCellFormatting = true,
                PdfUseWorksheetPrintSettings = true
            };
            var layout = ExcelConverter.ReadWorksheetPrintLayout(input, options);
            Assert.Equal(4, layout.Chunks.Count);
            Assert.True(layout.Landscape);
            Assert.Equal(0.85, layout.Scale, 3);
            Assert.Equal(0.35 * 72, layout.LeftMarginPoints, 3);
            Assert.Equal(0.45 * 72, layout.RightMarginPoints, 3);
            Assert.All(layout.Chunks, page => Assert.Single(page.ColumnWidths));
            Assert.Contains(layout.Chunks, page =>
                page.Rows.Any(r => r.Cells.Any(c => c.Text == "Student 8")));

            await ExcelConverter.ToPDFAsync(input, output, options);
            Assert.True(new FileInfo(output).Length > 100);
            var signature = new byte[5];
            using var pdf = File.OpenRead(output);
            Assert.Equal(5, pdf.Read(signature));
            Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(signature));
        }
        finally
        {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public async Task RepeatsPrintTitleFromOutsideThePrintArea()
    {
        var input = Temp(".xlsx");
        var output = Temp(".pdf");
        try
        {
            using (var book = new XLWorkbook())
            {
                var ws = book.AddWorksheet("Report");
                ws.Cell("A1").Value = "Report heading";
                ws.Cell("A2").Value = "Report subheading";
                ws.Cell("A4").Value = "First";
                ws.Cell("A5").Value = "Second";
                ws.Cell("A6").Value = "Third";
                ws.PageSetup.PrintAreas.Add("A4:A6");
                ws.PageSetup.SetRowsToRepeatAtTop(1, 2);
                ws.PageSetup.AddHorizontalPageBreak(5);
                book.SaveAs(input);
            }

            var pages = ExcelConverter.ReadWorksheetPrintLayout(input,
                new ExcelConversionOptions
                {
                    PdfUseWorksheetPrintSettings = true,
                    PdfPreserveCellFormatting = true
                }).Chunks;
            Assert.Equal(2, pages.Count);
            Assert.All(pages, p =>
            {
                Assert.Equal("Report heading", p.Header.Cells[0].Text);
                Assert.Equal("Report subheading", Assert.Single(p.AdditionalHeaders!).Cells[0].Text);
            });
            Assert.Equal("Third", pages[1].Rows[0].Cells[0].Text);
            await ExcelConverter.ToPDFAsync(input, output,
                new ExcelConversionOptions
                {
                    PdfUseWorksheetPrintSettings = true,
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

    [Fact]
    public async Task RepeatedPrintHeadersConsumeBudgetForEverySection()
    {
        using var input = new MemoryStream();
        using (var book = new XLWorkbook())
        {
            var sheet = book.AddWorksheet("Data");
            for (int i = 1; i <= 6; i++) sheet.Cell(i, 1).Value = "Row " + i;
            sheet.PageSetup.PrintAreas.Add("A1:A6");
            sheet.PageSetup.SetRowsToRepeatAtTop(1, 2);
            sheet.PageSetup.AddHorizontalPageBreak(4);
            book.SaveAs(input);
        }
        input.Position = 0;
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => ExcelConverter.ToPDFAsync(input, output,
            new ExcelConversionOptions
            {
                PdfUseWorksheetPrintSettings = true,
                PdfPreserveCellFormatting = true,
                PdfPrintScalePercent = 100,
                MaxSelectedCells = 6
            }));
        Assert.Equal(0, output.Length);
        Assert.True(input.CanRead);
        input.Position = 0;
        await ExcelConverter.ToPDFAsync(input, output, new ExcelConversionOptions
        {
            PdfUseWorksheetPrintSettings = true,
            PdfPreserveCellFormatting = true,
            PdfPrintScalePercent = 100,
            MaxSelectedCells = 8
        });
        Assert.True(output.Length > 0);
    }

    [Fact]
    public void RequiresOptInFormattingAndRejectsInvalidScale()
    {
        var input = Temp(".xlsx");
        try
        {
            using (var book = new XLWorkbook())
            {
                book.AddWorksheet("Data").Cell("A1").Value = "Title";
                book.SaveAs(input);
            }
            Assert.Throws<ArgumentException>(() =>
                ExcelConverter.ReadWorksheetPrintLayout(input, new ExcelConversionOptions
                {
                    PdfUseWorksheetPrintSettings = true
                }));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                ExcelConverter.ReadWorksheetPrintLayout(input, new ExcelConversionOptions
                {
                    PdfPreserveCellFormatting = true,
                    PdfUseWorksheetPrintSettings = true,
                    PdfPrintScalePercent = 500
                }));
        }
        finally { if (File.Exists(input)) File.Delete(input); }
    }

    [Fact]
    public void RejectsBreakCuttingThroughMergedCells()
    {
        var input = Temp(".xlsx");
        try
        {
            using (var book = new XLWorkbook())
            {
                var ws = book.AddWorksheet("Report");
                ws.Cell("A1").Value = "Header";
                ws.Cell("A2").Value = "Merged";
                ws.Range("A2:A4").Merge();
                ws.PageSetup.AddHorizontalPageBreak(3);
                book.SaveAs(input);
            }
            Assert.Throws<NotSupportedException>(() =>
                ExcelConverter.ReadWorksheetPrintLayout(input, new ExcelConversionOptions
                {
                    PdfPreserveCellFormatting = true,
                    PdfUseWorksheetPrintSettings = true
                }));
        }
        finally { if (File.Exists(input)) File.Delete(input); }
    }
}
