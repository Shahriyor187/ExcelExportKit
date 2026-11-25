# Export Studio — a live .NET export playground

A small, polished ASP.NET Core 10 app using the **published** [DataExport 2.1.0](https://www.nuget.org/packages/DataExport/2.1.0) and [ExcelExportKit 2.1.0](https://www.nuget.org/packages/ExcelExportKit/2.1.0) NuGet packages.

**No API keys, external databases, uploads or accounts are required.** Every download uses a small synthetic sales dataset (eight fictional orders). The site runs locally; no third-party scripts or fonts are loaded.

## Run in 60 seconds

Install the [.NET 10 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/10.0) and, from the ExcelExportKit repository root, run:

```powershell
dotnet run --project examples/ExportStudio/ExportStudio.csproj --urls http://localhost:5075
```

Visit **http://localhost:5075**. Click a format to generate and download a file.

## Three working examples

| Pipeline | NuGet package(s) | Download formats |
| --- | --- | --- |
| Direct C# records to file | DataExport | CSV, JSON, XML, XLSX |
| C# records → intermediate XLSX → converted file | DataExport + ExcelExportKit | CSV, JSON, PDF, DOCX |
| Styled sales/region workbook | DataExport | XLSX with two sheets, formulas, filtering and alternating rows |

### Minimal direct-export code

```csharp
using DataExportPackage;

var rows = new[] { new { Customer = "Amina", Score = 1450 } };
DataExporter.FromObjects(rows).ToExcel("students.xlsx",
    new DataExportOptions { ExcelIncludeHeaders = true });
```

### Excel → PDF code

```csharp
using ExportFiles.Services;

await ExcelConverter.ToPDFAsync("students.xlsx", "students.pdf",
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

## Endpoints

- `GET /api/health` — health and package labels
- `GET /api/preview` — eight fixed fictional orders
- `GET /api/export/data/{csv|json|xml|xlsx}` — DataExport direct output
- `GET /api/export/workbook/{csv|json|pdf|docx}` — both libraries in series
- `GET /api/export/report/xlsx` — a styled multi-sheet report

Each request creates a uniquely named directory under the OS temp directory, uses a fixed small dataset, reads the result into memory for download, and deletes the temporary files in `finally`.

## Safety / scope

This is a **sample**, not a production-ready upload or document-processing service. It deliberately has **no upload endpoint** and generates only a small fixed dataset. Production services need authentication, authorization, input limits, concurrency and execution limits, disk quotas, cleanup guarantees and monitoring. Some Excel print layout and graphic features are approximated by ExcelExportKit, not rendered pixel-for-pixel. The XLSX writer supports specific features, not a full Excel rendering engine.

The sample targets **.NET 10**. DataExport and ExcelExportKit also ship `net8.0` assemblies for compatibility, but the .NET 8 runtime is out of support.

## Tests

The [Export Studio smoke workflow](../../.github/workflows/export-studio.yml) restores the **public** NuGet packages, runs the server, downloads every supported format, checks file signatures/ZIP packages/JSON structure and verifies that unsupported routes are rejected. Existing library CI is separate.
