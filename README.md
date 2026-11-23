# ExcelExportKit

[![NuGet](https://img.shields.io/nuget/v/ExcelExportKit)](https://www.nuget.org/packages/ExcelExportKit)
[![Build](https://github.com/Shahriyor187/ExcelExportKit/actions/workflows/ci.yml/badge.svg)](https://github.com/Shahriyor187/ExcelExportKit/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE.txt)

A .NET library for converting Excel worksheets (`.xlsx`) to **CSV, JSON, PDF, or Word (DOCX)**, including worksheet selection and batch exports.

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

The formatting-aware PDF renderer now supports **fully selected horizontal and vertical merged cells**, solid per-edge borders, approximate Excel row heights, font sizes, colors and basic alignment. Merges crossing a selected range or from the repeated header into data are explicitly rejected. Dashed/double border styles are approximated with solid strokes. Embedded PNG/JPEG pictures are available separately as an opt-in, cell-anchored feature (below). Charts, exact print areas, manual page breaks, and pixel-perfect Excel pagination are not supported.

## Embed cell-anchored PNG/JPEG images in styled PDFs

Set `PdfIncludeEmbeddedImages = true` alongside `PdfPreserveCellFormatting = true`
to include a bounded number of PNG/JPEG drawings anchored at cell origins.
Images appear **inside the corresponding PDF table cells**, below the
cell's text, not at their original floating Excel coordinates.

```csharp
await ExcelConverter.ToPDFAsync("report.xlsx", "report.pdf",
    new ExcelConversionOptions
    {
        PdfPreserveCellFormatting = true,
        PdfIncludeEmbeddedImages = true,
        PdfMaxEmbeddedImages = 16,
        PdfMaxEmbeddedImageBytes = 2 * 1024 * 1024
    });
```

The default is **off**, preserving existing output. Images must be PNG/JPEG,
cell-anchored, no larger than 4,096 pixels per side, no larger than the
configured per-image byte limit, and within the selected range. Unsupported
floating pictures, offsets, merged-cell overlaps, multiple images per cell
and outside-range pictures produce explicit errors instead of being silently
ignored. Free-floating/pixel-perfect Excel graphics, charts, SVGs and
text-wrapped positioning are not yet supported.

## Worksheet print layout in PDF (opt-in)

Set `PdfUseWorksheetPrintSettings = true` together with
`PdfPreserveCellFormatting = true` to honor supported worksheet print
settings during PDF export:

- One or more rectangular Excel print areas (each rendered separately)
- A4 portrait or landscape orientation and page margins
- Workbook print scale (10–400%) or an optional explicit PDF scale override
- One repeated print title row, including a title above the selected print area
- Manual horizontal and vertical page breaks, splitting sections onto separate pages

```csharp
await ExcelConverter.ToPDFAsync("report.xlsx", "report.pdf",
    new ExcelConversionOptions
    {
        PdfPreserveCellFormatting = true,
        PdfUseWorksheetPrintSettings = true,
        PdfPrintScalePercent = 85
    });
```

This feature **does not reproduce the Microsoft Excel print engine**:
automatic table pagination can still add pages beyond manual breaks;
fit-to-N-pages, arbitrary paper sizes, multiple repeated title rows,
repeated title columns, print-centered placement, Excel header/footer text,
and exact page-boundary fidelity are not supported. Manual breaks crossing
merged cells are rejected. Unsupported ranges fail rather than silently
breaking merged-cell layout. Existing PDF output is unchanged unless opted in.

## PDF visual regression checks

The repository has a dedicated [PDF visual regression workflow](.github/workflows/pdf-visual.yml) that creates four reproducible spreadsheet/PDF pairs:

- `merged-header` — colored merged heading, border weights, custom column widths and alignment
- `vertical-merge` — two vertically merged groups with data beside row-spanning cells
- `multi-page` — 119 records to exercise PDF pagination and repeated headers
- `image-anchor` — visibly colored PNG and JPEG images anchored in different table cells

CI rasterizes each PDF with Poppler at 110 DPI, verifies expected text and pagination, checks that both colored embedded images are visible in rendered pixels, and uploads the source workbooks, PDFs, PNG pages, a JSON summary and per-page pixel-difference masks. On pull requests, it compares against the base branch **when the base contains the fixture generator**. The initial bootstrap PR therefore produces review artifacts without a comparison.

Run the fixtures locally:

```bash
dotnet run -c Release --project tests/ExcelExportKit.VisualChecks -- ./visual-output/candidate
python3 -m pip install Pillow==11.3.0
python3 tests/ExcelExportKit.VisualChecks/compare.py --candidate ./visual-output/candidate --output ./visual-output/report
```

The PNG stage requires Poppler's `pdftoppm` and `pdftotext` commands. Pixel differences **do not fail CI automatically**; intentional design changes should be inspected in the uploaded artifacts. This is regression review tooling, **not** a comparison to Excel's own printing engine or a guarantee of visual fidelity. PDF text checks are limited to fixture labels and do not replace human inspection.

### Comparing against independently exported Excel PDFs

Optionally, generate PDFs from the **same four fixture XLSX files** using desktop Microsoft Excel. Store them in a local reference folder as `merged-header.pdf`, `vertical-merge.pdf`, `multi-page.pdf`, and `image-anchor.pdf` (partial sets are allowed). Then run:

```bash
python3 tests/ExcelExportKit.VisualChecks/compare.py \
  --candidate ./visual-output/candidate \
  --excel-reference ./visual-output/excel-reference-pdfs \
  --output ./visual-output/excel-review
```

This produces `excelComparison` entries in the JSON report and difference masks in `excel-diffs/`. It intentionally **does not fail CI based on pixel differences**. Microsoft Excel print settings, fonts, graphics placement, and pagination can differ; reference PDFs must be produced independently and reviewed by a human. The repository does not currently contain validated Excel-generated reference PDFs. Do not label renderer-to-renderer comparisons as Excel fidelity.

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
