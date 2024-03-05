using ExportFiles.Services;

class Program
{
    static async Task Main(string[] args)
    {
        string inputFile = "C:\\Users\\Intel Computers\\Desktop\\test.xlsx";
        string jsonOutputFile = "C:\\Users\\Intel Computers\\Desktop\\Export\\file.json";
        try
        {
            await ExcelConverter.ToJsonAsync(inputFile, jsonOutputFile);
            Console.WriteLine("JSON conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error converting to JSON: " + ex.Message);
        }
        string pdfOutputFile = "C:\\Users\\Intel Computers\\Desktop\\Export\\file.pdf";
        try
        {
            await ExcelConverter.ToPDFAsync(inputFile, pdfOutputFile);
            Console.WriteLine("PDF conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error converting to PDF: " + ex.Message);
        }
        string wordfile = "C:\\Users\\Intel Computers\\Desktop\\Export\\file.docx";
        string csvfil = "C:\\Users\\Intel Computers\\Desktop\\Export\\file.csv";
        await ExcelConverter.ToWord(inputFile, wordfile);
        await ExcelConverter.ToCSVAsync(inputFile, csvfil);
        
    }
}