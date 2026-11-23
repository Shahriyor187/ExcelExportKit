using ClosedXML.Excel;
using ExportFiles.Services;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class WorkbookInputSafetyTests
{
    private static string Temp(string extension) =>
        Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

    private static string CreateWorkbook()
    {
        var input = Temp(".xlsx");
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Data");
        sheet.Cell("A1").Value = "Name";
        sheet.Cell("B1").Value = "Value";
        sheet.Cell("A2").Value = "Amina";
        sheet.Cell("B2").Value = 1450;
        wb.SaveAs(input);
        return input;
    }

    [Fact]
    public async Task EnforcesFileSizeBeforeReadingWorkbook()
    {
        var input = CreateWorkbook();
        var output = Temp(".json");
        try
        {
            var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelConverter.ToJsonAsync(input, output, new ExcelConversionOptions
                {
                    MaxWorkbookFileBytes = 64
                }));
            Assert.Contains("maximum", error.Message);
            Assert.False(File.Exists(output));
        }
        finally { File.Delete(input); if (File.Exists(output)) File.Delete(output); }
    }

    [Fact]
    public async Task EnforcesArchiveEntryAndExpandedSizeLimits()
    {
        var input = CreateWorkbook();
        var output = Temp(".csv");
        try
        {
            await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelConverter.ToCSVAsync(input, output, new ExcelConversionOptions
                {
                    MaxWorkbookArchiveEntries = 1
                }));
            await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelConverter.ToCSVAsync(input, output, new ExcelConversionOptions
                {
                    MaxUncompressedWorkbookBytes = 128
                }));
            Assert.False(File.Exists(output));
        }
        finally { File.Delete(input); if (File.Exists(output)) File.Delete(output); }
    }

    [Fact]
    public async Task RejectsExcessiveSelectedCellCountForJsonAndStyledPdf()
    {
        var input = CreateWorkbook();
        var json = Temp(".json");
        var pdf = Temp(".pdf");
        try
        {
            var options = new ExcelConversionOptions { MaxSelectedCells = 3 };
            await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelConverter.ToJsonAsync(input, json, options));
            options.PdfPreserveCellFormatting = true;
            await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelConverter.ToPDFAsync(input, pdf, options));
            Assert.False(File.Exists(json));
            Assert.False(File.Exists(pdf));
        }
        finally
        {
            File.Delete(input);
            if (File.Exists(json)) File.Delete(json);
            if (File.Exists(pdf)) File.Delete(pdf);
        }
    }

    [Fact]
    public async Task ValidLimitsDoNotChangeConversion()
    {
        var input = CreateWorkbook();
        var output = Temp(".json");
        try
        {
            await ExcelConverter.ToJsonAsync(input, output, new ExcelConversionOptions
            {
                MaxWorkbookFileBytes = 8 * 1024 * 1024,
                MaxUncompressedWorkbookBytes = 8 * 1024 * 1024,
                MaxWorkbookArchiveEntries = 100,
                MaxSelectedCells = 4,
                PreserveJsonTypes = true
            });
            Assert.Contains("Amina", await File.ReadAllTextAsync(output));
        }
        finally { File.Delete(input); if (File.Exists(output)) File.Delete(output); }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-10)]
    public async Task RejectsInvalidConfiguredLimits(long limit)
    {
        var input = CreateWorkbook();
        var output = Temp(".json");
        try
        {
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
                ExcelConverter.ToJsonAsync(input, output, new ExcelConversionOptions
                {
                    MaxWorkbookFileBytes = limit
                }));
        }
        finally { File.Delete(input); if (File.Exists(output)) File.Delete(output); }
    }

    [Fact]
    public void ValidatesPrintLayoutBeforeLoadingUntrustedWorkbook()
    {
        var input = CreateWorkbook();
        try
        {
            Assert.Throws<InvalidDataException>(() =>
                ExcelConverter.ReadWorksheetPrintLayout(input, new ExcelConversionOptions
                {
                    PdfPreserveCellFormatting = true,
                    PdfUseWorksheetPrintSettings = true,
                    MaxWorkbookFileBytes = 64
                }));
        }
        finally { File.Delete(input); }
    }
}
