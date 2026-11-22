using ClosedXML.Excel;
using ExportFiles.Services;
using System.Text;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class PdfMergedCellsTests
{
    [Fact]
    public async Task StyledPdfRendersMergedHeaderAndBorderedCells()
    {
        var input = Temp(".xlsx");
        var output = Temp(".pdf");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Report");
                sheet.Cell("A1").Value = "Combined heading";
                sheet.Range("A1:B1").Merge();
                sheet.Cell("C1").Value = "Score";
                sheet.Cell("A2").Value = "Amina";
                sheet.Cell("B2").Value = "Uzbekistan";
                sheet.Cell("C2").Value = 1490;
                sheet.Range("A2:C2").Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                sheet.Range("A1:B1").Style.Fill.BackgroundColor = XLColor.FromHtml("#224466");
                sheet.Row(2).Height = 33;
                sheet.Cell("C2").Style.Font.FontSize = 14;
                book.SaveAs(input);
            }

            var snapshot = ExcelConverter.ReadPdfFormatting(input, new ExcelConversionOptions
            {
                PdfPreserveCellFormatting = true
            });

            Assert.Equal(2, snapshot.Header.Cells.Count);
            Assert.Equal(2, snapshot.Header.Cells[0].ColumnSpan);
            Assert.Equal(1, snapshot.Header.Cells[0].RowSpan);
            Assert.Equal("#224466", snapshot.Header.Cells[0].Background);
            Assert.Equal(3, Assert.Single(snapshot.Rows).Cells.Count);
            var score = snapshot.Rows[0].Cells[2];
            Assert.Equal(14d, score.FontSize);
            Assert.True(score.VerticalPadding > 4);
            Assert.NotNull(score.RightBorder);

            await ExcelConverter.ToPDFAsync(input, output, new ExcelConversionOptions
            {
                PdfPreserveCellFormatting = true,
                DocumentTitle = "Results"
            });
            var pdf = File.ReadAllBytes(output);
            Assert.True(pdf.Length > 100);
            Assert.Equal("%PDF-", Encoding.ASCII.GetString(pdf, 0, 5));
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public async Task RowspanPreservesFollowingTableCellPositions()
    {
        var input = Temp(".xlsx");
        var output = Temp(".pdf");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Summary");
                sheet.Cell("A1").Value = "Group";
                sheet.Cell("B1").Value = "Name";
                sheet.Cell("A2").Value = "First group";
                sheet.Range("A2:A3").Merge();
                sheet.Cell("B2").Value = "Amina";
                sheet.Cell("B3").Value = "Ali";
                book.SaveAs(input);
            }
            var options = new ExcelConversionOptions
            {
                PdfPreserveCellFormatting = true,
                SkipEmptyRows = true
            };
            var result = ExcelConverter.ReadPdfFormatting(input, options);
            Assert.Equal(2, result.Rows.Count);
            Assert.Equal(2, result.Rows[0].Cells[0].RowSpan);
            Assert.Single(result.Rows[1].Cells);
            Assert.Equal("Ali", result.Rows[1].Cells[0].Text);
            await ExcelConverter.ToPDFAsync(input, output, options);
            Assert.True(new FileInfo(output).Length > 100);
        }
        finally { Delete(input, output); }
    }

    [Fact]
    public void RejectsCrossBoundaryMergeInSelectedRange()
    {
        var input = Temp(".xlsx");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Data");
                sheet.Cell("A1").Value = "Heading";
                sheet.Range("A1:C1").Merge();
                sheet.Cell("A2").Value = "Student";
                book.SaveAs(input);
            }
            var error = Assert.Throws<NotSupportedException>(() =>
                ExcelConverter.ReadPdfFormatting(input, new ExcelConversionOptions
                {
                    RangeAddress = "B1:C2",
                    PdfPreserveCellFormatting = true
                }));
            Assert.Contains("entire rectangle", error.Message);
        }
        finally { Delete(input); }
    }

    [Fact]
    public void RejectsMergeAcrossRepeatedHeaderBoundary()
    {
        var input = Temp(".xlsx");
        try
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("Data");
                sheet.Cell("A1").Value = "Heading";
                sheet.Range("A1:A2").Merge();
                sheet.Cell("B1").Value = "Score";
                sheet.Cell("B2").Value = 1400;
                book.SaveAs(input);
            }

            var error = Assert.Throws<NotSupportedException>(() =>
                ExcelConverter.ReadPdfFormatting(input, new ExcelConversionOptions
                {
                    PdfPreserveCellFormatting = true
                }));
            Assert.Contains("repeated PDF header", error.Message);
        }
        finally { Delete(input); }
    }

    private static string Temp(string ext) =>
        Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);

    private static void Delete(params string[] paths)
    {
        foreach (var path in paths)
            if (File.Exists(path)) File.Delete(path);
    }
}
