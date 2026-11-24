using ClosedXML.Excel;
using ExportFiles.Services;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class RowWindowTests
{
    private static void WithWorkbook(Action<string> test, bool merge = false)
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var path = Path.Combine(dir.FullName, "rows.xlsx");
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell(1, 1).Value = "Id";
                sheet.Cell(1, 2).Value = "Name";
                for (int i = 2; i <= 8; i++) { sheet.Cell(i, 1).Value = i - 1; sheet.Cell(i, 2).Value = "Row " + (i - 1); }
                if (merge) sheet.Range("B3:B4").Merge();
                workbook.SaveAs(path);
            }
            test(path);
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public void StandardExtractionPreservesHeadersAndTypes()
    {
        WithWorkbook(path =>
        {
            var result = ExcelConverter.ReadWorksheet(path, new ExcelConversionOptions { RowOffset = 2, MaxRows = 3, PreserveJsonTypes = true }, CancellationToken.None);
            Assert.Equal(new[] { "Id", "Name" }, result.Headers);
            Assert.Equal(new[] { "3", "4", "5" }, result.Rows.Select(row => row[0]));
            Assert.Equal(new[] { 3d, 4d, 5d }, result.TypedRows.Select(row => Assert.IsType<double>(row[0])));
        });
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(long.MaxValue, 3)]
    public void EmptyWindowRetainsHeader(long offset, long limit)
    {
        WithWorkbook(path =>
        {
            var options = new ExcelConversionOptions { RowOffset = offset, MaxRows = limit };
            var result = ExcelConverter.ReadWorksheet(path, options, CancellationToken.None);
            Assert.Empty(result.Rows);
            Assert.Equal("Id", result.Headers[0]);
            var pdf = ExcelConverter.ReadPdfFormatting(path, options);
            Assert.Empty(pdf.Rows);
            Assert.Equal("Id", pdf.Header.Cells[0].Text);
        });
    }

    [Fact]
    public void StyledPdfRejectsCroppedMergeButAllowsCompleteMerge()
    {
        WithWorkbook(path =>
        {
            Assert.Throws<NotSupportedException>(() => ExcelConverter.ReadPdfFormatting(path,
                new ExcelConversionOptions { RowOffset = 2, MaxRows = 1, PdfPreserveCellFormatting = true }));
            var result = ExcelConverter.ReadPdfFormatting(path,
                new ExcelConversionOptions { RowOffset = 1, MaxRows = 2, PdfPreserveCellFormatting = true });
            Assert.Equal(2, result.Rows.Count);
            Assert.Equal(2, result.Rows[0].Cells[1].RowSpan);
        }, merge: true);
    }

    [Fact]
    public void PrintLayoutRetainsOriginalHeader()
    {
        WithWorkbook(path =>
        {
            var layout = ExcelConverter.ReadWorksheetPrintLayout(path, new ExcelConversionOptions
            {
                RowOffset = 2, MaxRows = 3, PdfPreserveCellFormatting = true,
                PdfUseWorksheetPrintSettings = true, PdfPrintScalePercent = 100
            });
            var chunk = Assert.Single(layout.Chunks);
            Assert.Equal("Id", chunk.Header.Cells[0].Text);
            Assert.Equal(new[] { "3", "4", "5" }, chunk.Rows.Select(row => row.Cells[0].Text));
        });
    }

    [Theory]
    [InlineData(-1, 1)]
    [InlineData(0, -1)]
    public void NegativeBoundsAreRejected(long offset, long limit)
    {
        WithWorkbook(path => Assert.Throws<ArgumentOutOfRangeException>(() =>
            ExcelConverter.ReadWorksheet(path, new ExcelConversionOptions { RowOffset = offset, MaxRows = limit }, CancellationToken.None)));
    }
}
