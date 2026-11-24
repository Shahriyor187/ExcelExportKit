using System.Text;
using System.Text.Json;
using System.IO.Compression;
using DataExportPackage;
using ExportFiles.Services;

using var workbook = new MemoryStream();
var data = new[] { new { Id = 1, Text = "alpha" }, new { Id = 2, Text = "  =1+1" }, new { Id = 3, Text = "a\tb" } };
await DataExporter.FromObjects(data).ExportAsync(workbook, ExportFormat.Excel,
    new DataExportOptions { ExcelIncludeHeaders = true, ExcelAutoFitColumns = false });
Check(workbook.CanRead && workbook.CanWrite, "Workbook ownership");
foreach (var format in new[] { "csv", "tsv", "json", "pdf", "word" })
{
    workbook.Position = 0;
    using var output = new MemoryStream();
    var options = new ExcelConversionOptions { RowOffset = 1, MaxRows = 2, PreserveJsonTypes = true };
    switch (format)
    {
        case "csv": await ExcelConverter.ToCSVAsync(workbook, output, options); break;
        case "tsv": await ExcelConverter.ToTSVAsync(workbook, output, options); break;
        case "json": await ExcelConverter.ToJsonAsync(workbook, output, options); break;
        case "pdf": await ExcelConverter.ToPDFAsync(workbook, output, options); break;
        case "word": await ExcelConverter.ToWordAsync(workbook, output, options); break;
    }
    Check(workbook.CanRead && output.CanWrite, "Caller-owned streams");
    var bytes = output.ToArray();
    var text = Encoding.UTF8.GetString(bytes).Replace("\r\n", "\n");
    if (format == "csv") Check(text == "Id,Text\n2,'  =1+1\n3,a\tb\n", "CSV rows and protection");
    if (format == "tsv") Check(text == "Id\tText\n2\t'  =1+1\n3\t\"a\tb\"\n", "TSV rows and protection");
    if (format == "json")
    {
        using var json = JsonDocument.Parse(bytes);
        Check(json.RootElement.GetArrayLength() == 2 && json.RootElement[0].GetProperty("Id").GetInt32() == 2, "JSON row window and types");
    }
    if (format == "pdf") Check(text.StartsWith("%PDF-"), "PDF signature");
    if (format == "word")
    {
        output.Position = 0;
        using var zip = new ZipArchive(output, ZipArchiveMode.Read, true);
        Check(zip.GetEntry("word/document.xml") is not null, "DOCX structure");
    }
    Console.WriteLine("Cross-package format verified: " + format);
}

using var streamed = new MemoryStream();
await DataExporter.WriteExcelStreamingAsync(Rows(), streamed, new StreamingExcelOptions { Columns = new[] { "Id", "Text" } });
streamed.Position = 0;
using var csv = new MemoryStream();
await ExcelConverter.ToCSVAsync(streamed, csv);
Check(Encoding.UTF8.GetString(csv.ToArray()).Replace("\r\n", "\n") == "Id,Text\n1,streamed\n", "Forward-only XLSX interoperability");
Console.WriteLine("Cross-package integration passed.");

static async IAsyncEnumerable<object> Rows()
{
    yield return new { Id = 1, Text = "streamed" };
    await Task.CompletedTask;
}
static void Check(bool condition, string message)
{
    if (!condition) throw new InvalidDataException(message);
}
