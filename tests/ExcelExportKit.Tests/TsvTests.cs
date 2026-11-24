using ClosedXML.Excel;
using ExportFiles.Services;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class TsvTests
{
    private static async Task WithWorkbook(Func<string, string, Task> test)
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var path = Path.Combine(dir.FullName, "input.xlsx");
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell(1, 1).Value = "Id";
                sheet.Cell(1, 2).Value = "Text";
                sheet.Cell(2, 1).Value = 1;
                sheet.Cell(2, 2).Value = "skip";
                sheet.Cell(3, 1).Value = 2;
                sheet.Cell(3, 2).Value = "a\tb\"c\nnext";
                sheet.Cell(4, 1).Value = 3;
                sheet.Cell(4, 2).Value = "=1+1";
                var other = workbook.AddWorksheet("Other");
                other.Cell(1, 1).Value = "Name";
                other.Cell(2, 1).Value = "other";
                workbook.SaveAs(path);
            }
            await test(path, dir.FullName);
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public Task SelectedRowsAreQuotedAndOptionsAreUnchanged() => WithWorkbook(async (input, dir) =>
    {
        var output = Path.Combine(dir, "selected.tsv");
        var options = new ExcelConversionOptions { WorksheetName = "Data", RowOffset = 1, MaxRows = 2, CsvDelimiter = ';' };
        await ExcelConverter.ToTSVAsync(input, output, options);
        var nl = "\n";
        Assert.Equal("Id\tText" + nl + "2\t\"a\tb\"\"c\nnext\"" + nl + "3\t'=1+1" + nl, (await File.ReadAllTextAsync(output)).Replace("\r\n", "\n"));
        Assert.Equal(';', options.CsvDelimiter);
    });

    [Fact]
    public Task BatchAndAllSheetsUseTsvExtensionAndFormulaProtection() => WithWorkbook(async (input, dir) =>
    {
        var all = Path.Combine(dir, "all");
        var batch = Path.Combine(dir, "batch");
        var allResults = await ExcelConverter.ConvertAllSheetsAsync(input, all, "TSV");
        Assert.Equal(2, allResults.Count);
        Assert.All(allResults, r => Assert.True(r.Success, r.Error));
        Assert.Equal(2, Directory.GetFiles(all, "*.tsv").Length);
        Assert.Contains("3\t'=1+1", await File.ReadAllTextAsync(Path.Combine(all, "Data.tsv")));
        var result = Assert.Single(await ExcelConverter.ConvertBatchAsync(new[] { input }, batch, "tsv"));
        Assert.True(result.Success, result.Error);
        Assert.Contains("3\t'=1+1", await File.ReadAllTextAsync(Path.Combine(batch, "input.tsv")));
    });

    [Fact]
    public Task CancelledConversionPreservesExistingOutput() => WithWorkbook(async (input, dir) =>
    {
        var output = Path.Combine(dir, "output.tsv");
        await File.WriteAllTextAsync(output, "existing");
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelConverter.ToTSVAsync(input, output, cancellationToken: cancel.Token));
        Assert.Equal("existing", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(dir, ".export-*"));
    });

    [Fact]
    public Task ZeroRowsAndExplicitProtectionOptOut() => WithWorkbook(async (input, dir) =>
    {
        var output = Path.Combine(dir, "output.tsv");
        await ExcelConverter.ToTSVAsync(input, output, new ExcelConversionOptions { MaxRows = 0 });
        Assert.Equal("Id\tText" + Environment.NewLine, await File.ReadAllTextAsync(output));
        await ExcelConverter.ToTSVAsync(input, output, new ExcelConversionOptions { RowOffset = 2, MaxRows = 1, CsvProtectFormulas = false });
        Assert.Contains("3\t=1+1", await File.ReadAllTextAsync(output));
    });
}
