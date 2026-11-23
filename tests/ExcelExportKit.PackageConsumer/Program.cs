using ClosedXML.Excel;
using ExportFiles.Services;
using System.Text.Json;

var root = Path.Combine(Path.GetTempPath(), "ExcelExportKit-consumer-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    var source = Path.Combine(root, "sample.xlsx");
    using (var book = new XLWorkbook())
    {
        var ws = book.AddWorksheet("Scores");
        ws.Cell("A1").Value = "Student";
        ws.Cell("B1").Value = "Score";
        ws.Cell("A2").Value = "Amina";
        ws.Cell("B2").Value = 1450;
        ws.Cell("A3").Value = "Ali";
        ws.Cell("B3").Value = 1400;
        book.SaveAs(source);
    }

    var csv = Path.Combine(root, "result.csv");
    var json = Path.Combine(root, "result.json");
    var pdf = Path.Combine(root, "result.pdf");
    var docx = Path.Combine(root, "result.docx");
    var settings = new ExcelConversionOptions
    {
        WorksheetName = "Scores",
        PreserveJsonTypes = true,
        PdfPreserveCellFormatting = true,
        MaxWorkbookFileBytes = 4 * 1024 * 1024,
        MaxUncompressedWorkbookBytes = 8 * 1024 * 1024,
        MaxWorkbookArchiveEntries = 512,
        MaxSelectedCells = 100
    };
    await ExcelConverter.ToCSVAsync(source, csv, settings);
    await ExcelConverter.ToJsonAsync(source, json, settings);
    await ExcelConverter.ToPDFAsync(source, pdf, settings);
    await ExcelConverter.ToWordAsync(source, docx, settings);
    if (!File.ReadAllText(csv).Contains("Amina,1450", StringComparison.Ordinal))
        throw new Exception("CSV consumer output mismatch.");
    using (var data = JsonDocument.Parse(File.ReadAllText(json)))
    {
        if (data.RootElement.GetArrayLength() != 2 ||
            data.RootElement[0].GetProperty("Score").GetInt32() != 1450)
            throw new Exception("Typed JSON consumer output mismatch.");
    }
    var pdfHeader = new byte[5];
    using (var file = File.OpenRead(pdf))
    {
        if (file.Read(pdfHeader) != 5 ||
            System.Text.Encoding.ASCII.GetString(pdfHeader) != "%PDF-")
            throw new Exception("Invalid PDF package output.");
    }
    if (new FileInfo(docx).Length < 100)
        throw new Exception("DOCX package output is missing.");
    Console.WriteLine("Packed ExcelExportKit consumer OK: " +
        System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);
}
finally
{
    Directory.Delete(root, recursive: true);
}
