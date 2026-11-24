using ClosedXML.Excel;
using ExportFiles.Services;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace ExcelExportKit.Tests;

public sealed class StreamSecurityTests
{
    private static byte[] Workbook()
    {
        using var book = new XLWorkbook();
        var sheet = book.AddWorksheet("Data");
        sheet.Cell(1, 1).Value = "Id";
        sheet.Cell(1, 2).Value = "Text";
        sheet.Cell(2, 1).Value = 1;
        sheet.Cell(2, 2).Value = "=1+1";
        using var output = new MemoryStream();
        book.SaveAs(output);
        return output.ToArray();
    }

    [Theory]
    [InlineData("csv")]
    [InlineData("tsv")]
    [InlineData("json")]
    [InlineData("pdf")]
    [InlineData("styled")]
    [InlineData("print")]
    [InlineData("word")]
    public async Task StreamFormatsLeaveCallerResourcesOpen(string format)
    {
        using var input = new MemoryStream(Workbook(), writable: false);
        using var output = new MemoryStream();
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var options = new ExcelConversionOptions
            {
                TemporaryDirectory = dir.FullName,
                PdfPreserveCellFormatting = format is "styled" or "print",
                PdfUseWorksheetPrintSettings = format == "print",
                PdfPrintScalePercent = 100
            };
            await Convert(format, input, output, options);
            Assert.True(input.CanRead);
            Assert.True(output.CanWrite);
            Assert.NotEmpty(output.ToArray());
            Assert.Empty(dir.GetFileSystemInfos());
            Assert.Null(options.MaxWorkbookFileBytes);
            var text = Encoding.UTF8.GetString(output.ToArray());
            if (format == "csv") Assert.Contains("1,'=1+1", text);
            if (format == "tsv") Assert.Contains("1\t'=1+1", text);
            if (format is "pdf" or "styled" or "print") Assert.StartsWith("%PDF-", text);
            if (format == "word")
            {
                output.Position = 0;
                using var zip = new ZipArchive(output, ZipArchiveMode.Read, true);
                Assert.NotNull(zip.GetEntry("word/document.xml"));
            }
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public async Task DefaultRejectsNonSeekableWithoutReading()
    {
        using var input = new ForwardInput(Workbook());
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => ExcelConverter.ToCSVAsync(input, output));
        Assert.Equal(0, input.BytesRead);
        Assert.Equal(0, output.Length);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(InputBufferingMode.Memory)]
    [InlineData(InputBufferingMode.TemporaryFile)]
    public async Task ExplicitBufferingAcceptsExactLimitAndCleansUp(InputBufferingMode mode)
    {
        var bytes = Workbook();
        using var input = new ForwardInput(bytes);
        using var output = new MemoryStream();
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            await ExcelConverter.ToCSVAsync(input, output, new ExcelConversionOptions
            {
                InputBufferingMode = mode, MaxBufferedInputBytes = bytes.Length,
                TemporaryDirectory = dir.FullName
            });
            Assert.Equal(bytes.Length, input.BytesRead);
            Assert.True(input.CanRead);
            Assert.Contains("Id,Text", Encoding.UTF8.GetString(output.ToArray()));
            Assert.Empty(dir.GetFileSystemInfos());
        }
        finally { dir.Delete(true); }
    }

    [Theory]
    [InlineData(InputBufferingMode.Memory)]
    [InlineData(InputBufferingMode.TemporaryFile)]
    public async Task OversizedInputStopsAtLimitPlusOneAndCleansUp(InputBufferingMode mode)
    {
        using var input = new ForwardInput(new byte[8192]);
        using var output = new MemoryStream();
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            await Assert.ThrowsAsync<InvalidDataException>(() => ExcelConverter.ToCSVAsync(input, output,
                new ExcelConversionOptions { InputBufferingMode = mode, MaxBufferedInputBytes = 100,
                    MaxWorkbookFileBytes = 50, TemporaryDirectory = dir.FullName }));
            Assert.Equal(51, input.BytesRead);
            Assert.Equal(0, output.Length);
            Assert.True(input.CanRead);
            Assert.Empty(dir.GetFileSystemInfos());
        }
        finally { dir.Delete(true); }
    }

    [Theory]
    [InlineData("file")]
    [InlineData("expanded")]
    [InlineData("entries")]
    [InlineData("cells")]
    public async Task SeekableInputSecurityLimitsRejectBeforeOutput(string limit)
    {
        using var input = new MemoryStream(Workbook());
        using var output = new MemoryStream();
        var options = new ExcelConversionOptions
        {
            MaxWorkbookFileBytes = limit == "file" ? 1 : null,
            MaxUncompressedWorkbookBytes = limit == "expanded" ? 1 : null,
            MaxWorkbookArchiveEntries = limit == "entries" ? 1 : null,
            MaxSelectedCells = limit == "cells" ? 1 : null
        };
        await Assert.ThrowsAsync<InvalidDataException>(() => ExcelConverter.ToCSVAsync(input, output, options));
        Assert.Equal(0, output.Length);
        Assert.True(input.CanRead);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task CorruptZipDoesNotReplaceExistingDestination()
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var destination = Path.Combine(dir.FullName, "existing.csv");
            await File.WriteAllTextAsync(destination, "original");
            using var input = new ForwardInput(new byte[] { 1, 2, 3, 4 });
            await Assert.ThrowsAsync<InvalidDataException>(() => ExcelConverter.ToCSVAsync(input, destination,
                new ExcelConversionOptions { InputBufferingMode = InputBufferingMode.TemporaryFile,
                    TemporaryDirectory = dir.FullName }));
            Assert.Equal("original", await File.ReadAllTextAsync(destination));
            Assert.Single(dir.GetFileSystemInfos());
            Assert.True(input.CanRead);
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public async Task CancelDuringBufferingEvenIfSourceIgnoresToken()
    {
        using var cancel = new CancellationTokenSource();
        using var input = new ForwardInput(Workbook()) { AfterRead = () => cancel.Cancel() };
        using var output = new MemoryStream();
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var task = ExcelConverter.ToCSVAsync(input, output, new ExcelConversionOptions
            {
                InputBufferingMode = InputBufferingMode.TemporaryFile, TemporaryDirectory = dir.FullName
            }, cancel.Token);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
            Assert.True(task.IsCanceled);
            Assert.Equal(0, output.Length);
            Assert.True(input.CanRead);
            Assert.Empty(dir.GetFileSystemInfos());
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public async Task InvalidConfigurationAndAliasingRejectBeforeReading()
    {
        using var input = new ForwardInput(Workbook());
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => ExcelConverter.ToCSVAsync(input, output,
            new ExcelConversionOptions { InputBufferingMode = InputBufferingMode.Memory, MaxSelectedCells = 0 }));
        await Assert.ThrowsAsync<ArgumentException>(() => ExcelConverter.ToCSVAsync(input, input));
        Assert.Equal(0, input.BytesRead);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task MixedPathOverloadsProduceValidOutput()
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var source = Path.Combine(dir.FullName, "input.xlsx");
            var dest = Path.Combine(dir.FullName, "output.tsv");
            await File.WriteAllBytesAsync(source, Workbook());
            using var output = new MemoryStream();
            await ExcelConverter.ToCSVAsync(source, output);
            Assert.Contains("Id,Text", Encoding.UTF8.GetString(output.ToArray()));
            using var input = new MemoryStream(Workbook());
            await ExcelConverter.ToTSVAsync(input, dest);
            Assert.Contains("Id\tText", await File.ReadAllTextAsync(dest));
            Assert.True(input.CanRead);
            Assert.Equal(2, dir.GetFileSystemInfos().Length);
        }
        finally { dir.Delete(true); }
    }

    private static Task Convert(string format, Stream input, Stream output, ExcelConversionOptions options) =>
        format switch
        {
            "csv" => ExcelConverter.ToCSVAsync(input, output, options),
            "tsv" => ExcelConverter.ToTSVAsync(input, output, options),
            "json" => ExcelConverter.ToJsonAsync(input, output, options),
            "word" => ExcelConverter.ToWordAsync(input, output, options),
            _ => ExcelConverter.ToPDFAsync(input, output, options)
        };

    private sealed class ForwardInput(byte[] bytes) : MemoryStream(bytes)
    {
        public int BytesRead { get; private set; }
        public Action? AfterRead { get; init; }
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            var read = base.Read(buffer.Span);
            BytesRead += read;
            AfterRead?.Invoke();
            return ValueTask.FromResult(read);
        }
    }
}
