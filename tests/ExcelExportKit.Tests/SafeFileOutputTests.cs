using ExportFiles.Services;
using ClosedXML.Excel;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class SafeFileOutputTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
    private readonly string destination;
    private readonly string input;

    public SafeFileOutputTests()
    {
        Directory.CreateDirectory(directory);
        destination = Path.Combine(directory, "output.csv");
        File.WriteAllText(destination, "original");
        input = Path.Combine(directory, "input.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");
        sheet.Cell(1, 1).Value = "Name";
        sheet.Cell(2, 1).Value = "Amina";
        workbook.SaveAs(input);
    }

    private Task<ConversionFileResult> Write(long? limit = null, CancellationToken token = default) =>
        ExcelConverter.ConvertFileAsync(input, destination, "csv", maxOutputBytes: limit, cancellationToken: token);

    [Fact]
    public async Task SuccessfulPublicationReportsActualFileSize()
    {
        var result = await Write();
        Assert.Equal(destination, result.OutputPath);
        Assert.Equal(new FileInfo(destination).Length, result.OutputBytes);
        Assert.Contains("Amina", File.ReadAllText(destination));
        Assert.True(result.Elapsed >= TimeSpan.Zero);
        Assert.Empty(Directory.GetDirectories(directory));
    }

    [Fact]
    public async Task OutputLimitPreservesExistingFile()
    {
        await Assert.ThrowsAsync<IOException>(() => Write(1));
        Assert.Equal("original", File.ReadAllText(destination));
        Assert.Empty(Directory.GetDirectories(directory));
    }

    [Fact]
    public async Task ExactOutputLimitIsAccepted()
    {
        var first = await Write();
        var second = await Write(first.OutputBytes);
        Assert.Equal(first.OutputBytes, second.OutputBytes);
    }

    [Fact]
    public async Task InvalidLimitPreservesExistingFile()
    {
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => Write(0));
        Assert.Equal("original", File.ReadAllText(destination));
        Assert.Empty(Directory.GetDirectories(directory));
    }

    [Fact]
    public async Task CancellationPreservesExistingFile()
    {
        using var source = new CancellationTokenSource();
        source.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Write(token: source.Token));
        Assert.Equal("original", File.ReadAllText(destination));
        Assert.Empty(Directory.GetDirectories(directory));
    }

    [Fact]
    public async Task RenderFailurePreservesExistingFile()
    {
        File.Delete(input);
        await Assert.ThrowsAsync<FileNotFoundException>(() => Write());
        Assert.Equal("original", File.ReadAllText(destination));
        Assert.Empty(Directory.GetDirectories(directory));
    }
    
    [Fact]
    public async Task RejectsSameInputAndOutputPath()
    {
        var original = File.ReadAllBytes(input);
        await Assert.ThrowsAsync<ArgumentException>(async () =>
            await ExcelConverter.ConvertFileAsync(input, input, "csv"));
        Assert.Equal(original, File.ReadAllBytes(input));
    }

    public void Dispose() => Directory.Delete(directory, recursive: true);
}
