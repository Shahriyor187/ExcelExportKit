using ClosedXML.Excel;
using ExportFiles.Services;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class CancellationTests
{
    [Theory]
    [InlineData("csv")]
    [InlineData("json")]
    [InlineData("pdf")]
    [InlineData("docx")]
    [InlineData("formatted-pdf")]
    [InlineData("print-pdf")]
    public async Task PreCancelledConversionPreservesExistingOutput(string format)
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var output = Path.Combine(dir.FullName, "output." + format);
            await File.WriteAllTextAsync(output, "existing");
            using var cancel = new CancellationTokenSource();
            cancel.Cancel();
            var input = Path.Combine(dir.FullName, "not-opened.xlsx");
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                Convert(input, output, format, cancel.Token));
            Assert.Equal("existing", await File.ReadAllTextAsync(output));
            Assert.Single(dir.GetFiles());
        }
        finally { dir.Delete(true); }
    }

    [Theory]
    [InlineData("csv")]
    [InlineData("json")]
    [InlineData("pdf")]
    [InlineData("docx")]
    public async Task CancellableSuccessfulConversionReplacesOutputAndLeavesNoStagingFile(string format)
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var input = Path.Combine(dir.FullName, "input.xlsx");
            var output = Path.Combine(dir.FullName, "output." + format);
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell("A1").Value = "Id";
                sheet.Cell("A2").Value = 123;
                workbook.SaveAs(input);
            }
            await File.WriteAllTextAsync(output, "existing");
            using var cancel = new CancellationTokenSource();
            await Convert(input, output, format, cancel.Token);
            Assert.True(new FileInfo(output).Length > 0);
            Assert.NotEqual(System.Text.Encoding.UTF8.GetBytes("existing"), await File.ReadAllBytesAsync(output));
            Assert.Equal(2, dir.GetFiles().Length);
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public async Task PreCancelledBatchDoesNotCreateDirectoryOrEnumerateInput()
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var output = Path.Combine(dir.FullName, "output");
            var visited = false;
            IEnumerable<string> Inputs()
            {
                visited = true;
                yield return "missing.xlsx";
            }
            using var cancel = new CancellationTokenSource();
            cancel.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                ExcelConverter.ConvertBatchAsync(Inputs(), output, "csv", cancel.Token));
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                ExcelConverter.ConvertAllSheetsAsync("missing.xlsx", output, "csv", cancel.Token));
            Assert.False(visited);
            Assert.False(Directory.Exists(output));
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public async Task BatchCancellationAfterFirstItemPreservesCompletedOutput()
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var input = Path.Combine(dir.FullName, "input.xlsx");
            var output = Path.Combine(dir.FullName, "output");
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell("A1").Value = "Id";
                sheet.Cell("A2").Value = 123;
                workbook.SaveAs(input);
            }
            using var cancel = new CancellationTokenSource();
            var visited = 0;
            IEnumerable<string> Inputs()
            {
                visited++;
                yield return input;
                cancel.Cancel();
                visited++;
                yield return input;
                visited++;
                yield return input;
            }
            var task = ExcelConverter.ConvertBatchAsync(Inputs(), output, "csv", cancel.Token);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
            Assert.True(task.IsCanceled);
            Assert.Equal(2, visited);
            var file = Assert.Single(Directory.GetFiles(output));
            Assert.Equal("input.csv", Path.GetFileName(file));
            Assert.Contains("123", await File.ReadAllTextAsync(file));
        }
        finally { dir.Delete(true); }
    }

    private static Task Convert(string input, string output, string format, CancellationToken token)
    {
        var options = new ExcelConversionOptions
        {
            PdfPreserveCellFormatting = format is "formatted-pdf" or "print-pdf",
            PdfUseWorksheetPrintSettings = format == "print-pdf"
        };
        return format switch
        {
            "csv" => ExcelConverter.ToCSVAsync(input, output, options, token),
            "json" => ExcelConverter.ToJsonAsync(input, output, options, token),
            "docx" => ExcelConverter.ToWordAsync(input, output, options, token),
            _ => ExcelConverter.ToPDFAsync(input, output, options, token)
        };
    }
}
