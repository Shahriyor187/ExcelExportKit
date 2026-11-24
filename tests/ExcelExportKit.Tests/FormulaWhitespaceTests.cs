using ClosedXML.Excel;
using ExportFiles.Services;
using System.Text;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class FormulaWhitespaceTests
{
    [Theory]
    [InlineData("  =1+1", true)]
    [InlineData("\t+1", true)]
    [InlineData("\u00a0-1", true)]
    [InlineData(" \t@SUM(A1)", true)]
    [InlineData("  ordinary", false)]
    [InlineData("   ", false)]
    [InlineData("text=1+1", false)]
    public async Task DelimitedProtectionPreservesOriginalWhitespace(string value, bool formulaLike)
    {
        using var input = new MemoryStream();
        using (var book = new XLWorkbook())
        {
            var sheet = book.AddWorksheet("Data");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Text";
            sheet.Cell(2, 1).Value = 1;
            sheet.Cell(2, 2).Value = value;
            book.SaveAs(input);
        }
        foreach (var protect in new[] { true, false })
        foreach (var tsv in new[] { false, true })
        {
            using var output = new MemoryStream();
            input.Position = 0;
            var options = new ExcelConversionOptions { CsvProtectFormulas = protect };
            if (tsv) await ExcelConverter.ToTSVAsync(input, output, options);
            else await ExcelConverter.ToCSVAsync(input, output, options);
            var delimiter = tsv ? '\t' : ',';
            var field = protect && formulaLike ? "'" + value : value;
            if (field.Contains(delimiter))
                field = "\"" + field + "\"";
            var expected = $"Id{delimiter}Text{Environment.NewLine}1{delimiter}{field}{Environment.NewLine}";
            Assert.Equal(expected, Encoding.UTF8.GetString(output.ToArray()));
        }
    }
}
