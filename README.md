# ExcelExportKit

[![NuGet](https://img.shields.io/nuget/v/ExcelExportKit)](https://www.nuget.org/packages/ExcelExportKit)
[![Build](https://github.com/Shahriyor187/ExcelExportKit/actions/workflows/ci.yml/badge.svg)](https://github.com/Shahriyor187/ExcelExportKit/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE.txt)

A .NET library for converting Excel worksheets (`.xlsx`) to **CSV, JSON, PDF, or Word (DOCX)**, including worksheet selection and batch exports.

## Try Export Studio — interactive demo

**Want to see real output before integrating a package?** Run the [Export Studio demo](examples/ExportStudio/README.md), a responsive local ASP.NET Core application using both **ExcelExportKit 2.0.0** and **DataExport 2.0.0** from NuGet.

```powershell
dotnet run --project examples/ExportStudio/ExportStudio.csproj --urls http://localhost:5075
```

Open [http://localhost:5075](http://localhost:5075) to try CSV, JSON, XML, XLSX, styled PDF and DOCX downloads. No keys, accounts, upload endpoints or external services are required.

<details>
<summary>Copyable example: Excel worksheet → styled PDF</summary>

```csharp
using ExportFiles.Services;

await ExcelConverter.ToPDFAsync("sales.xlsx", "sales.pdf",
    new ExcelConversionOptions
    {
        PdfLandscape = true,
        PdfPreserveCellFormatting = true,
        MaxWorkbookFileBytes = 4 * 1024 * 1024,
        MaxUncompressedWorkbookBytes = 16 * 1024 * 1024,
        MaxWorkbookArchiveEntries = 256,
        MaxSelectedCells = 2000
    });
```

</details>



> **Versioning:** Install **ExcelExportKit 2.0.0 or newer** to use the v2 features described below. The GitHub source branch and the latest version actually available on NuGet.org can differ; check the installed package version rather than assuming a source branch is already published.
> **Runtime support:** The package still targets both `net8.0` and `net10.0` for compatibility. Microsoft .NET 8 runtime support ended on November 10, 2026. Prefer a supported .NET 10 runtime for production and review the hosting application's own lifecycle and security controls.


## Formatting-aware PDF (v2)

PDF conversion can now carry over **supported RGB cell backgrounds and text colors, bold/italic text, horizontal alignment and proportional Excel column widths**. It is opt-in to preserve the previous rendering behavior.

```csharp
await ExcelConverter.ToPDFAsync("report.xlsx", "report.pdf",
    new ExcelConversionOptions
    {
        PdfPreserveCellFormatting = true,
        PdfLandscape = true
    });
```

The formatting-aware PDF renderer now supports **fully selected horizontal and vertical merged cells**, solid per-edge borders, approximate Excel row heights, font sizes, colors and basic alignment. Merges crossing a selected range or from the repeated header into data are explicitly rejected. Dashed/double border styles are approximated with solid strokes. Embedded PNG/JPEG pictures are available separately as an opt-in, cell-anchored feature (below). Charts and pixel-perfect Excel pagination are not supported. Print areas and manual page breaks are available in the separate opt-in print-layout mode below.

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
- Repeated print title rows, including headings above the selected print area
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
fit-to-N-pages, arbitrary paper sizes,
repeated title columns, print-centered placement, Excel header/footer text,
and exact page-boundary fidelity are not supported. Manual breaks crossing
merged cells are rejected. Unsupported ranges fail rather than silently
breaking merged-cell layout. Existing PDF output is unchanged unless opted in.

## PDF visual regression checks

The repository has a dedicated [PDF visual regression workflow](.github/workflows/pdf-visual.yml) that creates five reproducible spreadsheet/PDF pairs:

- `merged-header` — colored merged heading, border weights, custom column widths and alignment
- `vertical-merge` — two vertically merged groups with data beside row-spanning cells
- `multi-page` — 119 records to exercise PDF pagination and repeated headers
- `image-anchor` — visibly colored PNG and JPEG images anchored in different table cells
- `print-layout` — opt-in landscape print area, 85% scaling, custom margins, two repeated title rows and both horizontal/vertical manual breaks

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

## Safeguards and performance validation

For untrusted XLSX input, options-based conversion supports optional limits
for the compressed input size, expanded XLSX archive bytes, archive entry
count and selected cell count. **All are disabled by default** to preserve
existing conversion behavior. Set your own limits based on the application's
upload and memory budgets.

```csharp
await ExcelConverter.ToJsonAsync("uploaded.xlsx", "safe.json",
    new ExcelConversionOptions
    {
        MaxWorkbookFileBytes = 20 * 1024 * 1024,
        MaxUncompressedWorkbookBytes = 100 * 1024 * 1024,
        MaxWorkbookArchiveEntries = 2048,
        MaxSelectedCells = 250_000
    });
```

These preflight checks do not turn ClosedXML into a streaming or sandboxed
parser. Web applications should also enforce upload limits, authentication,
rate limits, concurrency and timeouts.

A dedicated, manually triggered [Excel conversion benchmark workflow](.github/workflows/benchmarks.yml)
can profile synthetic 1K, 10K and 100K-row workbooks. It uploads a JSON report
with times and memory observations; 100K runs benchmark CSV/JSON only. Results
are **not** performance guarantees or proof of an optimization.

The release gates and known limitations are in
[the v2 release checklist](docs/RELEASE_V2_CHECKLIST.md).

## Validate a representative workbook locally before v2 release

The release candidate includes a **local-only** probe for checking that a
realistic workbook can be converted to CSV, JSON, PDF and DOCX with opt-in
resource limits. It deletes temporary conversions after checking them and
reports timings, output sizes, success/failure categories and a cumulative
process peak-memory figure, without printing worksheet contents.

```bash
dotnet run -c Release --project tests/ExcelExportKit.ReleaseProbe -- \
  ./local-report.xlsx --formats csv,json,pdf,docx --report ./release-probe.json
```

Run this in a local source checkout. Do **not** commit or upload confidential
workbooks. The probe provides diagnostic evidence, not a proof of exact
document fidelity or a complete security assessment. For the full protocol,
see [the v2 release checklist](docs/RELEASE_V2_CHECKLIST.md).

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
- By default, conversion uses the first worksheet. Options support named/indexed worksheets and batch conversion; they do not merge sheets into a single output document.
- JSON remains string-oriented by default; the `PreserveJsonTypes` option emits native numbers and booleans and ISO date strings.
- PDF can preserve supported styling, cell-anchored PNG/JPEG pictures and selected print-layout features when enabled, but does not support exact Excel printing, charts or floating drawings. Word output is primarily tabular.
- Spreadsheet formulas are not recalculated as part of conversion; values are taken from the workbook reader.

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

## Memory and cancellation characteristics (v2.1 development)

Cancellation-aware file conversion overloads accept an explicit `CancellationToken` after options. Cancellation is cooperative during archive checks, worksheet extraction and output row construction. ClosedXML workbook loading and third-party PDF/DOCX serialization may not be interruptible. Cancellable file exports stage output beside the destination, then move it into place after conversion and a final cancellation check. Completed batch items remain on disk; cancellation propagates rather than returning a success result. ClosedXML loads workbook data into memory; `MaxWorkbookFileBytes` limits compressed input bytes and `MaxSelectedCells` limits the selected conversion rectangle, not total process memory. CSV and JSON avoid accumulating an additional complete output string but still load the workbook and selected rows in memory.


## Partial worksheet conversions (v2.1 development)

Set `RowOffset` and `MaxRows` on `ExcelConversionOptions` to select physical data rows after the original header, before SkipEmptyRows filtering. Null MaxRows means unlimited; zero means no data rows. Negative values throw ArgumentOutOfRangeException. A window beyond the available rows retains the original header. Styled PDF conversion rejects merged rectangles cut by the row window and omits pictures in excluded data rows. Print-layout PDF counts data rows across print regions in their existing order, before horizontal page splitting; repeated print titles do not count. Existing input safety limits still apply to the original selected range and workbook. ClosedXML still loads the workbook into memory; a row window is not a streaming workbook reader.


## TSV conversion (v2.1 development)

Use `ExcelConverter.ToTSVAsync(inputFile, outputFile, options, cancellationToken)` to convert a worksheet, or pass `"tsv"` to `ConvertAllSheetsAsync` and `ConvertBatchAsync`. TSV forces a tab delimiter without modifying the supplied options, retains worksheet and range selection and row windows, and uses UTF-8 without BOM. Fields containing tabs, quotes or newlines use CSV-style quoting and doubled embedded quotes. Formula protection is enabled by default, including batch conversion. Cancellable path conversions retain the existing staged-output behavior. Use a parser that understands quoted TSV.


## Stream input and output (v2.1 development)

CSV, TSV, JSON, PDF and Word support stream-to-stream, path-to-stream and stream-to-path overloads. The caller owns input and output streams. Seekable input must start at position zero; seekable output must be empty and positioned at zero. Non-seekable output is supported. Input and output must not refer to the same underlying resource (distinct wrappers cannot always be detected). Do not change streams or options concurrently during conversion.

New stream overloads default to 64 MiB compressed workbook size, 256 MiB expanded ZIP size, 4096 archive entries and 1,000,000 selected cells. Supply positive custom limits to adjust them. Legacy path-to-path defaults are unchanged. Non-seekable input is rejected unless InputBufferingMode.Memory or InputBufferingMode.TemporaryFile is explicitly selected. Buffering enforces the smaller of MaxBufferedInputBytes (64 MiB by default) and MaxWorkbookFileBytes while copying. CSV/TSV/JSON write directly; PDF and Word stage rendered output in a temporary directory. Stream-to-path writes stage on the destination filesystem before replacement.

TemporaryDirectory must be an existing, application-controlled directory; never derive it or output paths directly from untrusted upload names. Unix temporary files/directories use owner-only permissions; Windows relies on the parent directory ACL, so use a private application directory. Temporary resources are cleaned up during normal completion, failure or cooperative cancellation, not guaranteed after process termination. Use external cleanup and disk quotas as appropriate.

ClosedXML still loads the workbook into memory. Limits are not a sandbox or a guaranteed process memory ceiling. For hostile uploads use isolated workers, OS memory/CPU/disk quotas, request limits and timeouts. Cancellation cannot interrupt every synchronous parser or renderer call. Caller-owned output may contain partial bytes on failure or cancellation; discard it. Spreadsheet formula protection is enabled by default for CSV/TSV stream conversions but is not a universal guarantee across spreadsheet applications.
