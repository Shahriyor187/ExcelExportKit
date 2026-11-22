# ExcelExportKit

[![NuGet](https://img.shields.io/nuget/v/ExcelExportKit)](https://www.nuget.org/packages/ExcelExportKit)
[![Build](https://github.com/Shahriyor187/ExcelExportKit/actions/workflows/ci.yml/badge.svg)](https://github.com/Shahriyor187/ExcelExportKit/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A .NET library for converting the first worksheet in an Excel workbook (`.xlsx`) to **CSV, JSON, PDF, or Word (DOCX)**.

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

- Only the **first worksheet** is converted.
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

This project is licensed under the MIT License. See [LICENSE](LICENSE).

## Limitations

- The library reads `.xlsx` files, not legacy binary `.xls` files.
- It converts the first worksheet only; it does not merge multiple worksheets.
- Values are exported as formatted text, not as strongly typed JSON numbers or dates.
- PDF and Word output focus on worksheet data; complex Excel styling, charts, images, and formulas are not reproduced as spreadsheet features.
