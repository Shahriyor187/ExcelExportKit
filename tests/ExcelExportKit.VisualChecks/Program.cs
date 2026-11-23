using ClosedXML.Excel;
using ExportFiles.Services;

if (args.Length != 1 || string.IsNullOrWhiteSpace(args[0]))
{
    Console.Error.WriteLine("Usage: dotnet run -- <output-directory>");
    return 2;
}

var outputDirectory = Path.GetFullPath(args[0]);
Directory.CreateDirectory(outputDirectory);

await RenderFixture("merged-header", book =>
{
    var sheet = book.AddWorksheet("Sales");
    sheet.Cell("A1").Value = "Q4 Sales Overview";
    sheet.Range("A1:C1").Merge();
    sheet.Range("A1:C1").Style.Fill.BackgroundColor = XLColor.FromHtml("#25466A");
    sheet.Range("A1:C1").Style.Font.FontColor = XLColor.White;
    sheet.Range("A1:C1").Style.Font.Bold = true;
    sheet.Cell("A2").Value = "Region";
    sheet.Cell("B2").Value = "Product";
    sheet.Cell("C2").Value = "Revenue";
    sheet.Cell("A3").Value = "Tashkent";
    sheet.Cell("B3").Value = "Subscriptions";
    sheet.Cell("C3").Value = 250000;
    sheet.Cell("A4").Value = "Samarkand";
    sheet.Cell("B4").Value = "Exports";
    sheet.Cell("C4").Value = 175000;
    sheet.Range("A2:C4").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
    sheet.Range("A2:C4").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
    sheet.Column(1).Width = 24;
    sheet.Column(2).Width = 22;
    sheet.Column(3).Width = 16;
    sheet.Row(3).Height = 28;
    sheet.Cell("C3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
});

await RenderFixture("vertical-merge", book =>
{
    var sheet = book.AddWorksheet("Students");
    sheet.Cell("A1").Value = "Group";
    sheet.Cell("B1").Value = "Student";
    sheet.Cell("C1").Value = "SAT Score";
    sheet.Range("A1:C1").Style.Fill.BackgroundColor = XLColor.FromHtml("#2C4E73");
    sheet.Range("A1:C1").Style.Font.FontColor = XLColor.White;
    sheet.Cell("A2").Value = "Foundation";
    sheet.Range("A2:A3").Merge();
    sheet.Cell("B2").Value = "Amina";
    sheet.Cell("C2").Value = 1500;
    sheet.Cell("B3").Value = "Ali";
    sheet.Cell("C3").Value = 1420;
    sheet.Cell("A4").Value = "Bachelor";
    sheet.Range("A4:A5").Merge();
    sheet.Cell("B4").Value = "Sardor";
    sheet.Cell("C4").Value = 1490;
    sheet.Cell("B5").Value = "Laylo";
    sheet.Cell("C5").Value = 1550;
    sheet.Range("A2:C5").Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
    sheet.Column(1).Width = 21;
    sheet.Column(2).Width = 25;
    sheet.Column(3).Width = 15;
});

await RenderFixture("multi-page", book =>
{
    var sheet = book.AddWorksheet("Results");
    sheet.Cell("A1").Value = "Student";
    sheet.Cell("B1").Value = "Score";
    sheet.Cell("C1").Value = "Status";
    sheet.Range("A1:C1").Style.Fill.BackgroundColor = XLColor.FromHtml("#25466A");
    sheet.Range("A1:C1").Style.Font.FontColor = XLColor.White;
    sheet.Column(1).Width = 32;
    sheet.Column(2).Width = 16;
    sheet.Column(3).Width = 22;
    for (var i = 2; i <= 120; i++)
    {
        sheet.Cell(i, 1).Value = "Student " + (i - 1);
        sheet.Cell(i, 2).Value = 900 + (i * 7) % 700;
        sheet.Cell(i, 3).Value = i % 3 == 0 ? "Needs review" : "Completed";
        if (i % 2 == 0)
            sheet.Range(i, 1, i, 3).Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F5F9");
    }
});

await RenderFixture("image-anchor", book =>
{
    var sheet = book.AddWorksheet("Student Cards");
    sheet.Cell("A1").Value = "Student";
    sheet.Cell("B1").Value = "Image";
    sheet.Cell("A2").Value = "Amina";
    sheet.Cell("A3").Value = "Ali";
    sheet.Column(1).Width = 24;
    sheet.Column(2).Width = 25;
    sheet.Row(2).Height = 60;
    sheet.Range("A1:B1").Style.Fill.BackgroundColor = XLColor.FromHtml("#25466A");
    sheet.Range("A1:B1").Style.Font.FontColor = XLColor.White;

    // Distinctive solid colors allow CI to prove that both formats are actually
    // visible after PDF rasterization (a 1x1 transparent image cannot do that).
    var magentaPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAADAAAAAgCAIAAADbtmxLAAAAP0lEQVR4nO3OMQHAIADAMEDD/EvBECb27+nJjkRB5nn2+JN1O/AlVISKUBEqQkWoCBWhIlSEilARKkJFqAiVF9ZiAf0XmBgdAAAAAElFTkSuQmCC");
    var greenJpeg = Convert.FromBase64String(
        "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wAARCAAgADADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDPooor+Kz/ADXCiiigAooooAKKKKACiiigAooooA//2Q==");
    using var png = new MemoryStream(magentaPng);
    using var jpeg = new MemoryStream(greenJpeg);
    sheet.AddPicture(png, ClosedXML.Excel.Drawings.XLPictureFormat.Png).MoveTo(sheet.Cell("B2"));
    sheet.AddPicture(jpeg, ClosedXML.Excel.Drawings.XLPictureFormat.Jpeg).MoveTo(sheet.Cell("B3"));
});

await RenderFixture("print-layout", book =>
{
    var sheet = book.AddWorksheet("Print Layout");
    sheet.Cell("A1").Value = "Section";
    sheet.Cell("B1").Value = "Item";
    sheet.Cell("C1").Value = "Score";
    sheet.Cell("D1").Value = "Status";
    sheet.Cell("A2").Value = "Print section";
    sheet.Cell("B2").Value = "Report title row";
    sheet.Cell("C2").Value = "Scale 85%";
    sheet.Cell("D2").Value = "Landscape";
    sheet.Range("A1:D1").Style.Fill.BackgroundColor = XLColor.FromHtml("#25466A");
    sheet.Range("A1:D1").Style.Font.FontColor = XLColor.White;
    sheet.Range("A1:D2").Style.Font.Bold = true;
    sheet.Range("A2:D2").Style.Fill.BackgroundColor = XLColor.FromHtml("#DCEAF5");

    for (var i = 1; i <= 32; i++)
    {
        var row = i + 2;
        sheet.Cell(row, 1).Value = i <= 16 ? "Part A" : "Part B";
        sheet.Cell(row, 2).Value = "Item " + i;
        sheet.Cell(row, 3).Value = 1100 + i * 9;
        sheet.Cell(row, 4).Value = i % 4 == 0 ? "Review" : "Complete";
        if (i % 2 == 0)
            sheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F0F5FA");
    }

    sheet.Column(1).Width = 19;
    sheet.Column(2).Width = 25;
    sheet.Column(3).Width = 16;
    sheet.Column(4).Width = 23;
    sheet.PageSetup.PrintAreas.Add("A1:D34");
    sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
    sheet.PageSetup.Scale = 85;
    sheet.PageSetup.Margins.Top = 0.55;
    sheet.PageSetup.Margins.Bottom = 0.60;
    sheet.PageSetup.Margins.Left = 0.40;
    sheet.PageSetup.Margins.Right = 0.45;
    sheet.PageSetup.SetRowsToRepeatAtTop(1, 2);
    sheet.PageSetup.AddHorizontalPageBreak(18);
    sheet.PageSetup.AddVerticalPageBreak(2);
});

Console.WriteLine("Created five XLSX/PDF fixture pairs in " + outputDirectory);
return 0;

async Task RenderFixture(string name, Action<XLWorkbook> create)
{
    var xlsx = Path.Combine(outputDirectory, name + ".xlsx");
    var pdf = Path.Combine(outputDirectory, name + ".pdf");
    using (var workbook = new XLWorkbook())
    {
        create(workbook);
        workbook.SaveAs(xlsx);
    }
    await ExcelConverter.ToPDFAsync(xlsx, pdf, new ExcelConversionOptions
    {
        DocumentTitle = "Visual fixture: " + name,
        PdfPreserveCellFormatting = true,
        PdfIncludeEmbeddedImages = name == "image-anchor",
        PdfLandscape = name == "merged-header",
        PdfUseWorksheetPrintSettings = name == "print-layout"
    });
    var signature = new byte[5];
    await using var file = File.OpenRead(pdf);
    if (await file.ReadAsync(signature) != 5 || System.Text.Encoding.ASCII.GetString(signature) != "%PDF-")
        throw new InvalidDataException("Expected PDF output for fixture " + name);
    Console.WriteLine(name + ": " + new FileInfo(pdf).Length + " bytes");
}
