using ClosedXML.Excel;
using ExportFiles.Services;
using System.Text.Json;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class StreamingOutputRegressionTests
{
    private static string Temp(string extension) =>
        Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

    [Fact]
    public async Task CsvStreamingPreservesEscapesFormulaProtectionAndUtf8()
    {
        var source = Temp(".xlsx");
        var target = Temp(".csv");
        try
        {
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.AddWorksheet("Data");
                sheet.Cell("A1").Value = "Name";
                sheet.Cell("B1").Value = "Note";
                sheet.Cell("A2").Value = "Amina, \"M\"";
                sheet.Cell("B2").Value = "=SUM(A1:A2)";
                sheet.Cell("A3").Value = "Toshkent";
                sheet.Cell("B3").Value = "Line 1\nLine 2";
                wb.SaveAs(source);
            }
            await ExcelConverter.ToCSVAsync(source, target, new ExcelConversionOptions());
            var text = await File.ReadAllTextAsync(target);
            Assert.StartsWith("Name,Note", text);
            Assert.Contains("\"Amina, \"\"M\"\"\",'=SUM(A1:A2)", text);
            Assert.Contains("\"Line 1\nLine 2\"", text);
            var bytes = await File.ReadAllBytesAsync(target);
            Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF);
        }
        finally { if(File.Exists(source)) File.Delete(source); if(File.Exists(target)) File.Delete(target); }
    }

    [Fact]
    public async Task TypedAndLegacyStreamingJsonPreservePublicSemantics()
    {
        var source = Temp(".xlsx");
        var typed = Temp(".json");
        var legacy = Temp(".json");
        try
        {
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.AddWorksheet("Data");
                sheet.Cell("A1").Value = "Score";
                sheet.Cell("B1").Value = "Passed";
                sheet.Cell("C1").Value = "Note";
                sheet.Cell("A2").Value = 1450;
                sheet.Cell("B2").Value = true;
                sheet.Cell("C2").Value = "Hello, \"world\"\nToshkent";
                sheet.Cell("A3").Value = 1300;
                sheet.Cell("B3").Value = false;
                wb.SaveAs(source);
            }
            await ExcelConverter.ToJsonAsync(source, typed,
                new ExcelConversionOptions { PreserveJsonTypes = true });
            await ExcelConverter.ToJsonAsync(source, legacy, new ExcelConversionOptions());

            using var typedDoc = JsonDocument.Parse(await File.ReadAllTextAsync(typed));
            using var legacyDoc = JsonDocument.Parse(await File.ReadAllTextAsync(legacy));
            Assert.Equal(2, typedDoc.RootElement.GetArrayLength());
            Assert.Equal(1450, typedDoc.RootElement[0].GetProperty("Score").GetDouble());
            Assert.True(typedDoc.RootElement[0].GetProperty("Passed").GetBoolean());
            Assert.Contains("Toshkent", typedDoc.RootElement[0].GetProperty("Note").GetString());
            Assert.Equal("1450", legacyDoc.RootElement[0].GetProperty("Score").GetString());
            Assert.Equal("True", legacyDoc.RootElement[0].GetProperty("Passed").GetString());
            Assert.Equal("", legacyDoc.RootElement[1].GetProperty("Note").GetString());
        }
        finally
        {
            foreach (var path in new[] { source, typed, legacy })
                if (File.Exists(path)) File.Delete(path);
        }
    }
}
