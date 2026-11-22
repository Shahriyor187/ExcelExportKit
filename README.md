# ExcelExportKit

[![NuGet](https://img.shields.io/nuget/v/ExcelExportKit)](https://www.nuget.org/packages/ExcelExportKit)
[![Build](https://github.com/Shahriyor187/ExcelExportKit/actions/workflows/ci.yml/badge.svg)](https://github.com/Shahriyor187/ExcelExportKit/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE.txt)

A .NET library for converting the first worksheet in an Excel workbook (`.xlsx`) to **CSV, JSON, PDF, or Word (DOCX)**.

## Formatting-aware PDF (v2 foundation)

PDF conversion can now carry over **supported RGB cell backgrounds and text colors, bold/italic text, horizontal alignment and proportional Excel column widths**. It is opt-in to preserve the previous rendering behavior.

```csharp
await ExcelConverter.ToPDFAsync("report.xlsx", "report.pdf",
    new ExcelConversionOptions
    {
        PdfPreserveCellFormatting = true,
        PdfLandscape = true
    });
```

The formatting-aware PDF renderer now supports **fully selected horizontal and vertical merged cells**, solid per-edge borders, approximate Excel row heights, font sizes, colors and basic alignment. Merges crossing a selected range or from the repeated header into data are explicitly rejected. Dashed/double border styles are approximated with solid strokes. It still does **not** reproduce embedded images/charts, exact print areas, manual page breaks, or pixel-perfect Excel pagination.

## Extract only the cells you need

Select a named sheet and an **A1-style range** for CSV, JSON, PDF, or Word conversion. The first row of the range is treated as its column headers; blank rows can be excluded.

```csharp
await ExcelConverter.ToJsonAsync("results.xlsx", "results.json",
    new ExcelConversionOptions
    {
        WorksheetName = "Scores",
        RangeAddress = "B3:D500",
        SkipEmptyRows = true,
        PreserveJsonTypes = true
    });
```

This is an in-memory ClosedXML reader. It does not yet provide forward-only extraction for extremely large uploaded workbooks.

## Requirements

- .NET 8 or .NET 10

## Installation

```bash
dotnet add package ExcelExportKit
```

## Quick start

```csharp
using ExportFiles.Services;

await ExcelConverter.ToCSVAsync("input.xlsx", "output.csv");
await ExcelConverter.ToJsonAsync("input.xlsx", "output.json");
await ExcelConverter.ToPDFAsync("input.xlsx", "output.pdf");
await ExcelConverter.ToWordAsync("input.xlsx", "output.docx");
```

The original `ToWord(inputFile, outputFile)` method remains available for compatibility.

## Conversion behavior

- Legacy methods convert the **first worksheet**; overloads accept a worksheet name or zero-based index.
- The first used row is treated as the header row.
- Blank headers are named `Column1`, `Column2`, and so on.
- Duplicate headers are made unique by appending a suffix (for example, `Name_2`).
- Cell values are converted to their formatted display text using invariant culture.
- Empty cells are exported as empty strings.
- CSV fields containing commas, double quotes, or line breaks are quoted and escaped.
- JSON output is an array of objects, using the header cells as property names.
- PDF output is a paginated A4 document.
- Word output contains one table, with the first row formatted as a header.

## Dependencies and licensing

- [ClosedXML](https://github.com/ClosedXML/ClosedXML) — MIT, for reading XLSX workbooks.
- [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) — MIT, for creating DOCX files.
- [TerraPDF](https://terrapdf.com/) — MIT, for generating PDF files.

This project is licensed under the MIT License. See [LICENSE](LICENSE.txt).

## Limitations

- The library reads `.xlsx` files, not legacy binary `.xls` files.
- It converts the first worksheet only; it does not merge multiple worksheets.
- Values are exported as formatted text, not as strongly typed JSON numbers or dates.
- PDF and Word output focus on worksheet data; complex Excel styling, charts, images, and formulas are not reproduced as spreadsheet features.

## New in 1.2.0

### Select a worksheet

```csharp
await ExcelConverter.ToJsonAsync("report.xlsx", "sales.json",
    new ExcelConversionOptions { WorksheetName = "Sales", PreserveJsonTypes = true });

await ExcelConverter.ToCSVAsync("report.xlsx", "sheet2.csv",
    new ExcelConversionOptions { WorksheetIndex = 1, CsvDelimiter = ';' });
```

For typed JSON, numeric cells become JSON numbers, Boolean cells become booleans, and
dates become JSON date strings. The existing two-argument `ToJsonAsync` retains
its string-based JSON format.

### Tables in PDF and customizable Word output

PDF exports now render table cells with column alignment, repeating headers
and page numbers instead of concatenating row text with pipes.

```csharp
await ExcelConverter.ToPDFAsync("report.xlsx", "report.pdf",
    new ExcelConversionOptions { DocumentTitle = "Sales", PdfLandscape = true });

await ExcelConverter.ToWordAsync("report.xlsx", "report.docx",
    new ExcelConversionOptions
    {
        DocumentTitle = "Sales",
        WordTableStyle = "TableGrid",
        WordLandscape = true,
        WordHeaderText = "Sales report",
        WordFooterText = "Internal",
        WordColumnWidthTwips = 1800
    });
```

### Batch and all-worksheet conversion

```csharp
var sheetResults = await ExcelConverter.ConvertAllSheetsAsync(
    "report.xlsx", "output/sheets", "csv");

var files = new[] { "january.xlsx", "february.xlsx" };
var batchResults = await ExcelConverter.ConvertBatchAsync(
    files, "output/months", "pdf");

foreach (var result in batchResults)
{
    if (!result.Success) Console.Error.WriteLine(result.Error);
}
```

These methods return a result per workbook or worksheet. A failed conversion
doesn't stop the remaining entries. Supported formats are `csv`, `json`, `pdf`
and `docx`. Output directory creation is automatic.

### ASP.NET Core example

```csharp
app.MapGet("/report", async () =>
{
    var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
    try
    {
        await ExcelConverter.ToPDFAsync("sales.xlsx", output,
            new ExcelConversionOptions { DocumentTitle = "Sales report" });
        return Results.File(await File.ReadAllBytesAsync(output),
            "application/pdf", "sales.pdf");
    }
    finally
    {
        if (File.Exists(output)) File.Delete(output);
    }
});
```

The options-based CSV overload protects formula-like values by default. The original two-argument `ToCSVAsync` preserves its prior output for compatibility.

The library is independent of ASP.NET Core. Validate any user-supplied input
paths and file sizes before opening Excel workbooks. PDF export currently uses
the built-in Latin font set; additional scripts may require custom fonts.
