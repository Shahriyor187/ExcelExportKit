using ClosedXML.Excel;
using ExportFiles.Services;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class ProgressTests
{
    private sealed class Observer(Action<ExportProgress>? action = null) : IProgress<ExportProgress>
    {
        public List<ExportProgress> Reports { get; } = new();
        public void Report(ExportProgress value) { Reports.Add(value); action?.Invoke(value); }
    }

    private static MemoryStream Workbook()
    {
        var output = new MemoryStream();
        using var book = new XLWorkbook();
        var sheet = book.AddWorksheet("Data");
        sheet.Cell(1, 1).Value = "Id";
        for (int i = 2; i <= 9; i++) sheet.Cell(i, 1).Value = i - 1;
        book.SaveAs(output);
        output.Position = 0;
        return output;
    }

    private static Task Convert(string format, Stream input, Stream output, ExcelConversionOptions options, CancellationToken token = default) =>
        format switch
        {
            "csv" => ExcelConverter.ToCSVAsync(input, output, options, token),
            "tsv" => ExcelConverter.ToTSVAsync(input, output, options, token),
            "json" => ExcelConverter.ToJsonAsync(input, output, options, token),
            "word" => ExcelConverter.ToWordAsync(input, output, options, token),
            _ => ExcelConverter.ToPDFAsync(input, output, options, token)
        };

    [Theory]
    [InlineData("csv")]
    [InlineData("tsv")]
    [InlineData("json")]
    [InlineData("word")]
    [InlineData("pdf")]
    [InlineData("styled")]
    [InlineData("print")]
    public async Task ReportsSelectedRowsAcrossFormats(string format)
    {
        using var input = Workbook();
        using var output = new MemoryStream();
        var observer = new Observer();
        await Convert(format, input, output, new ExcelConversionOptions
        {
            Progress = observer, ProgressIntervalRows = 2, RowOffset = 2, MaxRows = 5,
            PreserveJsonTypes = true, PdfPreserveCellFormatting = format is "styled" or "print",
            PdfUseWorksheetPrintSettings = format == "print", PdfPrintScalePercent = 100
        });
        AssertReports(observer, new long[] { 2, 4, 5 }, 5);
        Assert.True(input.CanRead);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task ZeroRowsReportZero()
    {
        using var input = Workbook();
        using var output = new MemoryStream();
        var observer = new Observer();
        await ExcelConverter.ToCSVAsync(input, output, new ExcelConversionOptions
        { Progress = observer, MaxRows = 0 });
        AssertReports(observer, new long[] { 0 }, 0);
    }

    [Fact]
    public async Task ExactIntervalDoesNotDuplicateFinalNotification()
    {
        using var input = Workbook();
        using var output = new MemoryStream();
        var observer = new Observer();
        await ExcelConverter.ToJsonAsync(input, output, new ExcelConversionOptions
        { Progress = observer, ProgressIntervalRows = 2, MaxRows = 4 });
        AssertReports(observer, new long[] { 2, 4 }, 4);
    }

    [Theory]
    [InlineData("csv")]
    [InlineData("word")]
    [InlineData("pdf")]
    public async Task CancellationInsideCallbackStopsAndCleansUp(string format)
    {
        using var input = Workbook();
        using var output = new MemoryStream();
        using var cancel = new CancellationTokenSource();
        var observer = new Observer(_ => cancel.Cancel());
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Convert(format, input, output,
                new ExcelConversionOptions { Progress = observer, ProgressIntervalRows = 2,
                    TemporaryDirectory = dir.FullName }, cancel.Token));
            AssertReports(observer, new long[] { 2 }, 8);
            Assert.True(input.CanRead);
            Assert.True(output.CanWrite);
            Assert.Empty(dir.GetFileSystemInfos());
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public async Task CallbackFailurePreservesExistingFileWithoutCancellationToken()
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var source = Path.Combine(dir.FullName, "input.xlsx");
            var destination = Path.Combine(dir.FullName, "output.csv");
            using var input = Workbook();
            await File.WriteAllBytesAsync(source, input.ToArray());
            await File.WriteAllTextAsync(destination, "original");
            var observer = new Observer(_ => throw new InvalidOperationException("callback failure"));
            var error = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                ExcelConverter.ToCSVAsync(source, destination,
                    new ExcelConversionOptions { Progress = observer, ProgressIntervalRows = 2 }));
            Assert.Equal("callback failure", error.Message);
            Assert.Equal("original", await File.ReadAllTextAsync(destination));
            AssertReports(observer, new long[] { 2 }, 8);
            Assert.Equal(2, dir.GetFileSystemInfos().Length);
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public async Task ConcurrentExportsHaveIndependentCounts()
    {
        using var a = Workbook();
        using var b = Workbook();
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        var x = new Observer();
        var y = new Observer();
        await Task.WhenAll(
            ExcelConverter.ToCSVAsync(a, first, new ExcelConversionOptions { Progress = x, ProgressIntervalRows = 2, MaxRows = 5 }),
            ExcelConverter.ToJsonAsync(b, second, new ExcelConversionOptions { Progress = y, ProgressIntervalRows = 3, MaxRows = 7 }));
        AssertReports(x, new long[] { 2, 4, 5 }, 5);
        AssertReports(y, new long[] { 3, 6, 7 }, 7);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public async Task InvalidIntervalRejectedBeforeOutput(int interval)
    {
        using var input = Workbook();
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => ExcelConverter.ToCSVAsync(input, output,
            new ExcelConversionOptions { ProgressIntervalRows = interval }));
        Assert.Equal(0, output.Length);
        Assert.Equal(0, input.Position);
    }

    private static void AssertReports(Observer observer, long[] rows, long total)
    {
        Assert.Equal(rows, observer.Reports.Select(p => p.RowsWritten));
        Assert.All(observer.Reports, p => { Assert.Equal((long?)total, p.TotalRows); Assert.True(p.Elapsed >= TimeSpan.Zero); });
        for (int i = 1; i < observer.Reports.Count; i++)
            Assert.True(observer.Reports[i].Elapsed >= observer.Reports[i - 1].Elapsed);
    }
}
