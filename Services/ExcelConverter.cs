using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using System.Text;

namespace ExportFiles.Services;
public static class ExcelConverter
{
    static ExcelConverter()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    public static async Task ToCSVAsync(string inputFile, string outputFile)
    {
        try
        {
            using (var stream = new FileStream(inputFile, FileMode.Open))
            {
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        throw new InvalidDataException("The Excel file does not contain any worksheets");
                    }

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    StringBuilder csvString = new StringBuilder();

                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            if (col > 1)
                                csvString.Append(",");
                            var cellValue = worksheet.Cells[row, col].Value;
                            if (cellValue != null)
                            {
                                string cellText = cellValue.ToString().Replace("\"", "\"\"");
                                if (cellText.Contains(",") || cellText.Contains("\"") ||
                                    cellText.Contains("\r") || cellText.Contains("\n"))
                                {
                                    cellText = $"\"{cellText}\"";
                                }
                                csvString.Append(cellText);
                            }
                        }
                        csvString.AppendLine();
                    }

                    await File.WriteAllTextAsync(outputFile, csvString.ToString(), Encoding.UTF8);
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to CSV: " + ex.Message);
        }
    }

    public static async Task ToJsonAsync(string inputFile, string outputFile)
    {
        try
        {
            using (var stream = new FileStream(inputFile, FileMode.Open))
            {
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        throw new InvalidDataException("The Excel file does not contain any worksheets");
                    }
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var dataTable = new DataTable();
                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        if (row == 1)
                        {
                            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                            {
                                dataTable.Columns.Add(worksheet.Cells[row, col].Value?.ToString() ?? "");
                            }
                        }
                        else
                        {
                            var dataRow = dataTable.NewRow();
                            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                            {
                                dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? "";
                            }
                            dataTable.Rows.Add(dataRow);
                        }
                    }
                    if (dataTable.Rows.Count == 0)
                    {
                        throw new InvalidDataException("The Excel file does not contain any data");
                    }
                    string jsonOutput = JsonConvert.SerializeObject(dataTable, Newtonsoft.Json.Formatting.Indented);
                    await File.WriteAllTextAsync(outputFile, jsonOutput, Encoding.UTF8);
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to JSON: " + ex.Message);
        }
    }

    public static async Task ToPDFAsync(string inputFile, string outputFile)
    {
        try
        {
            using (var stream = new FileStream(inputFile, FileMode.Open))
            {
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        throw new InvalidDataException("The Excel file does not contain any worksheets");
                    }

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    using (var document = new iTextSharp.text.Document())
                    {
                        PdfWriter.GetInstance(document, new FileStream(outputFile, FileMode.Create));
                        document.Open();

                        PdfPTable table = new PdfPTable(worksheet.Dimension.Columns);
                        for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                        {
                            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Value;
                                table.AddCell(cellValue != null ? cellValue.ToString() : "");
                            }
                        }

                        document.Add(table);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to PDF: " + ex.Message);
        }
    }

    public static async Task ToWord(string inputFile, string outputFile)
    {
        try
        {
            using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(stream))
            {
                var worksheet = excelPackage.Workbook.Worksheets[0];

                using (var wordDocument = WordprocessingDocument.Create(outputFile, WordprocessingDocumentType.Document))
                {
                    var mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    var body = mainPart.Document.AppendChild(new Body());

                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        var tableRow = new TableRow();

                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                            var tableCell = new TableCell(new Paragraph(new Run(new Text(cellValue))));
                            tableRow.Append(tableCell);
                        }

                        body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table(tableRow));
                    }
                }
            }

            Console.WriteLine("Conversion to Word completed successfully.");
        }
        catch (Exception ex)
        {
            throw new Exception("Error converting Excel to Word: " + ex.Message);
        }
    }
}