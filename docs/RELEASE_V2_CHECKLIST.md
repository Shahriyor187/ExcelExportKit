# ExcelExportKit v2 release readiness

This checklist prepares the package for **v2.0.0**. It is not a declaration
that v2 is shipped, and must not be used as evidence that NuGet contains the
current master branch. The repository currently keeps its declared package
version at 1.2.0; changing it and publishing are separate release steps.

## Shipping gates

- [ ] All .NET 8 and .NET 10 unit/integration tests pass on the release commit.
- [ ] The PDF visual comparison workflow passes on the release commit; manually
      review differences, because pixel changes intentionally do not fail CI.
- [ ] Produce independent reference PDFs using Microsoft Excel for representative
      merged cells, print-layout, images and multi-page reports; record test system
      fonts, Excel version and print settings. Compare using
      \`--excel-reference\` and review the diff images.
- [ ] Run the manual benchmark workflow at 1K, 10K and 100K rows; save JSON
      reports and examine wall time, peak working set, and output sizes.
      The 100K scenario benchmarks CSV/JSON only (not PDF/Word).
- [ ] Benchmark at least one actual customer workbook with permission, using
      no secret or personal data in CI artifacts.
- [ ] Review security/zip-input limits for server usage and make sure upstream
      upload limits, request timeouts, concurrency controls, and storage quotas
      are enforced by the hosting application.
- [ ] Inspect any new security advisories and known limitations for ClosedXML,
      OpenXML SDK and TerraPDF. No automated vulnerability audit is a substitute.
- [ ] Review public API compatibility and confirm the intended major version.
- [ ] Change the version to 2.0.0 in ExcelExportKit.csproj and adjust
      .github/workflows/publish.yml manual workflow input default accordingly.
- [ ] Update CHANGELOG.md release date and replace this pending checklist.
- [ ] Build the .nupkg/.snupkg, inspect the README, license and DLLs for both
      target frameworks; test a clean consumer using *only* the local package
      source before release.
- [ ] Obtain approval for publishing v2.0.0 to NuGet.org, then trigger the
      manual trusted-publishing workflow and verify the public NuGet listing.
- [ ] Create a GitHub tag/release after publication and publish release notes.

## Opt-in safeguards for untrusted XLSX

The package now exposes four nullable limits on \`ExcelConversionOptions\`:

| Setting | Example | What it restricts |
| --- | --- | --- |
| \`MaxWorkbookFileBytes\` | \`20 * 1024 * 1024\` | XLSX input file size |
| \`MaxUncompressedWorkbookBytes\` | \`100 * 1024 * 1024\` | Total expanded ZIP entry bytes |
| \`MaxWorkbookArchiveEntries\` | \`2048\` | XLSX ZIP entry count |
| \`MaxSelectedCells\` | \`250_000\` | Number of cells in selected range |

**All limits default to null** for compatibility. These guardrails can reject
obviously oversized workbooks before ClosedXML allocates their full object
model, but are not a streaming parser or a complete security sandbox.
In particular, they do not enforce CPU time, graph/object allocations,
filesystem limits, server concurrency or malicious-content safety. Add
application-level upload and execution limits for untrusted content.

\`\`\`csharp
var options = new ExcelConversionOptions
{
    WorksheetName = "Report",
    MaxWorkbookFileBytes = 20 * 1024 * 1024,
    MaxUncompressedWorkbookBytes = 100 * 1024 * 1024,
    MaxWorkbookArchiveEntries = 2048,
    MaxSelectedCells = 250_000,
    PdfPreserveCellFormatting = true
};
await ExcelConverter.ToPDFAsync("uploaded.xlsx", "report.pdf", options);
\`\`\`

Use options-based conversions to enable these limits. The legacy two-argument
and existing multi-file/all-sheets helpers are unchanged and do not expose
the safety-limit options.

## Benchmarks

Run locally on .NET 10:

\`\`\`bash
dotnet run -c Release --project tests/ExcelExportKit.Benchmarks -- 1000 ./benchmarks/1000
dotnet run -c Release --project tests/ExcelExportKit.Benchmarks -- 10000 ./benchmarks/10000
dotnet run -c Release --project tests/ExcelExportKit.Benchmarks -- 100000 ./benchmarks/100000
\`\`\`

Each produces a deterministic six-column XLSX, conversion artifacts and
\`benchmark.json\`. The 100K-run intentionally excludes PDF and DOCX because
full in-memory layout is not yet bounded. The harness reports elapsed time,
output size, current managed heap and peak process working set; these are
observational metrics, not performance guarantees. Use the same OS/runner,
.NET version and resources when comparing revisions.

The GitHub Actions manual workflow **Excel conversion benchmarks** exposes
row-size choices and retains benchmark reports as artifacts.

## Explicit remaining gaps

- Print-layout is an approximation of Excel, not exact rendering.
- The library uses in-memory ClosedXML and can still consume substantial memory
  for highly styled or complex worksheets.
- Excel chart rendering, floating drawings, fit-to-page, and arbitrary paper
  sizes are not implemented.
- Visual comparison against Microsoft Excel has optional tooling, but a curated
  set of validated reference PDFs is not yet committed.
- The benchmark infrastructure exists, but measured results require running
  the harness. Do not claim improvements without before/after measurements.
