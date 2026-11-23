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
- [x] Run the isolated benchmark workflow at 1K, 10K and 100K rows,
      review elapsed time and per-format process peak working set.
      All three runs passed on 2026-11-23.
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

## Validated isolated benchmark results (2026-11-23)

Each result is the median of three warmed-up, isolated conversions on a GitHub-hosted
.NET 10 runner, using a synthetic six-column XLSX. Peak resident memory
includes the worker's warmup, but not fixture generation or other formats.

| Format | 1K median ms | 10K median ms | 100K median ms |
| --- | ---: | ---: | ---: |
| CSV | 170.06 | 641.87 | 2044.29 |
| JSON | 203.63 | 743.27 | 2480.57 |
| PDF | 446.46 | 1125.34 | Not run |
| DOCX | 247.01 | 931.16 | Not run |

| Format | 1K peak RSS MiB | 10K peak RSS MiB | 100K peak RSS MiB |
| --- | ---: | ---: | ---: |
| CSV | 99.2 | 154.1 | 424.3 |
| JSON | 103.1 | 156.4 | 414.4 |
| PDF | 108.1 | 231.4 | Not run |
| DOCX | 103.9 | 169.7 | Not run |

Runs: [1K](https://github.com/Shahriyor187/ExcelExportKit/actions/runs/45146058345),
[10K](https://github.com/Shahriyor187/ExcelExportKit/actions/runs/45146513792),
[100K](https://github.com/Shahriyor187/ExcelExportKit/actions/runs/45146702987).
These measurements cannot be compared directly with the older cold,
single-process figures or generalized to arbitrary documents, production
concurrency, or hardware. A real workbook and independent Excel PDF
references remain outstanding.

## Updated measurement protocol (post-baseline)

The first benchmark series (1K, 10K, 100K rows) used a single process for
workbook generation and every format. Its peak working set was cumulative and
not attributable to a particular conversion. The initial 100K run measured
around 626 MiB of *whole-process* peak working set, not CSV/JSON alone.

The new `run_benchmarks.py` protocol generates the workbook separately. Each
format runs in a **new worker process** for each of three trials. That worker
performs one untimed same-format warmup, runs a full GC, then measures one
conversion. The report includes per-trial elapsed milliseconds, whole-process
allocations, current working set and process peak working set, plus medians.
Peak working set includes the warmup in that worker but excludes workbook
generation or other output formats. All metrics remain environment-dependent.

Use the manual GitHub Actions benchmark workflow at 1K, 10K, and 100K rows
to capture **new** baseline results before making improvement claims. These
results are not directly comparable to the previous cold, single-process
benchmark. At 100K the workflow still limits itself to CSV and JSON. It is
not a streaming XLSX reader, and concurrent server conversions may still
require explicit resource caps.

## Memory optimization notes

CSV and JSON output paths now write incrementally, avoiding one large CSV
StringBuilder or a second full JSON records list and serialized JSON string.
The worksheet reader no longer allocates typed value rows when native types
are not requested. It still materializes formatted display rows with ClosedXML
and therefore remains memory-intensive for large spreadsheets.

## Benchmarks

Run the new protocol locally with .NET 10:

```bash
dotnet build -c Release tests/ExcelExportKit.Benchmarks/ExcelExportKit.Benchmarks.csproj
python3 tests/ExcelExportKit.Benchmarks/run_benchmarks.py \
  --dll tests/ExcelExportKit.Benchmarks/bin/Release/net10.0/ExcelExportKit.Benchmarks.dll \
  --rows 10000 --output ./benchmark-results --repeats 3
```

Other supported sizes are 1,000 and 100,000 rows. The JSON report includes
per-trial metrics and medians; output artifacts are removed between trials.
The GitHub Actions manual workflow uses the same program.

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
