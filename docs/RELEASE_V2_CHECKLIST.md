# ExcelExportKit v2 release readiness

This checklist documents final review for the **v2.0.0** release branch. It is
not a declaration that v2 has shipped and must not be used as evidence that
NuGet contains this release branch. The stable package version is staged here,
but the stable v2.0.0 package has **not** been published.

## Shipping gates

- [ ] All .NET 8 and .NET 10 unit/integration tests pass on the release commit.
- [ ] The PDF visual comparison workflow passes on the release commit; manually
      review differences, because pixel changes intentionally do not fail CI.
- [x] Review independent Excel LTSC reference PDFs for four existing visual
      fixtures (merged cells, images, multi-page report). All expected labels
      and all 119 report records match, while pagination differs. See
      [the reference review](EXCEL_REFERENCE_FIDELITY_REVIEW_2026-11-23.md).
- [x] Independently review the opt-in print-settings mode against Excel LTSC:
      A4 landscape, four-page horizontal/vertical breaks, 32 records and
      repeated titles match. Page geometry and actual 85% scaling still differ;
      this is **approximate** rendering, not Excel parity. See
      [advanced print-layout review](EXCEL_PRINT_LAYOUT_REFERENCE_2026-11-23.md).

- [x] Run the isolated benchmark workflow at 1K, 10K and 100K rows,
      review elapsed time and per-format process peak working set.
      All three runs passed on 2026-11-23.
      The 100K scenario benchmarks CSV/JSON only (not PDF/Word).
- [x] Smoke-test an independently published, non-private real-world Excel
      workbook with the local release probe: UK ONS
      `summaryofrecords2026q3fqe.xlsx` on Windows .NET 10.0.15.
      CSV/JSON/PDF/DOCX structural conversions succeeded (2026-11-23).
      The local probe selected the first worksheet by default, validates basic
      output structure, not all workbook sheets or per-cell fidelity. Workbook
      content and converted outputs were not committed.
- [x] Add a synthetic production-like multi-sheet invoice integration test
      spanning CSV, typed JSON, styled PDF and DOCX with Unicode, numeric and
      Boolean fields. This supplements but does not replace the real-workbook gate.
- [ ] Verify the consuming server's upload limits, request timeouts,
      concurrency controls, process isolation, and storage quotas. The library
      only provides opt-in XLSX preflight limits; it cannot enforce hosting
      safeguards inside an unknown consuming application.
- [ ] Manually inspect current advisories, licenses and known limitations for
      ClosedXML, OpenXML SDK and TerraPDF before production publication. CI's
      automated direct/transitive vulnerability audit passed for the previous
      release-candidate commit, but cannot replace this review.
- [ ] Final manual review of public API compatibility and upgrade semantics,
      including two-argument CSV behavior versus options-based defaults. The
      intended package release remains major version 2.0.0.
- [x] Prepare the `2.0.0-rc.1` candidate version, leaving NuGet publishing disabled.
- [x] Stage stable `2.0.0` metadata and changelog in a separate branch,
      **without** publishing; run CI and visual validation on the exact commit
      before merge or release approval.
- [x] Draft the 2026-11-23 CHANGELOG.md v2.0.0 release notes;
      publication remains pending.
- [ ] Build the candidate .nupkg/.snupkg, inspect README, license, both DLLs
      and package version. Test clean .NET 8 and .NET 10 consumers against the
      **locally built candidate** (with a package-source mapping that prevents
      silent fallback to a public version).
- [ ] Complete direct/transitive NuGet advisory checks, reviewing any findings
      or audit-feed access failures before release.
- [ ] Obtain explicit approval to publish stable v2.0.0 to NuGet.org, then
      trigger the guarded manual trusted-publishing workflow from the reviewed
      stable release commit and verify the public NuGet listing.
- [ ] Create a GitHub tag/release after publication and publish release notes.

## Release-candidate procedure

1. CI must pass on the candidate branch: .NET 8/10 tests, package audit,
   local NuGet consumer installations and PDF visual regression.
2. Inspect the generated `ExcelExportKit.2.0.0-rc.1.nupkg` and `.snupkg`
   artifacts from Actions; this is verification, **not** publication.
3. Export the same visual-fixture `.xlsx` workbooks independently using
   Microsoft Excel to PDFs, record Excel version, print settings and fonts,
   and review rendered image differences against the library output.
4. Verify representative real-world workbooks and server-side resource limits.
5. Only after all gates and explicit release approval, create a separate
   stable-version change and execute the protected NuGet publishing workflow.

The publishing workflow requests the exact version and a `PUBLISH`
confirmation; it is **manual**. A passed CI workflow is not authorization
to publish.

## Privacy-conscious real-workbook release probe

A local-only command-line probe is available under
`tests/ExcelExportKit.ReleaseProbe` to help validate a representative workbook
without uploading its contents. It converts a selected sheet using the public
options API and enforces conservative preflight limits (20 MiB file,
100 MiB expanded ZIP, 2,048 entries and 250,000 selected cells).

Example (Windows PowerShell; run from a local repository checkout):

```powershell
dotnet run -c Release --project tests/ExcelExportKit.ReleaseProbe -- "C:\\Work\\sample.xlsx" --sheet "Report" --formats csv,json,pdf,docx --report release-probe.json
```

The probe deletes its converted output files after checking that they were
created. It logs **only** each format's success, elapsed time, output size,
exception *type*, and a cumulative process peak-memory measurement. It does
not print workbook paths, worksheet values, sheet names, or exception messages.
The optional report is local unless the user explicitly shares it; no report
or workbook is automatically uploaded anywhere.

**Important:** This is not a complete integrity or security audit. It checks
file signatures for PDF/DOCX and nonempty outputs, not visual appearance or
cell-by-cell equivalence. It processes input in-memory through ClosedXML and
must not be used on confidential or untrusted workbooks without appropriate
local safeguards. If real data are sensitive, use a synthetic or anonymized
equivalent and inspect the four outputs locally before sharing any diagnostics.

The 2026-11-23 Excel LTSC comparisons confirmed synthetic fixture content
and print-section order but found layout differences, including scaling and
pagination. Approximate Excel printing remains an explicitly documented
limitation for stable v2.

## 2026-11-23 real-world smoke-test evidence

A user downloaded the public UK Office for National Statistics workbook
`summaryofrecords2026q3fqe.xlsx` and ran
`ExcelExportKit.ReleaseProbe` locally on Windows with .NET 10.0.15.
Only the sanitized JSON report was shared, not the workbook or converted outputs.

| Format | Result | Elapsed time | Output bytes |
| --- | --- | ---: | ---: |
| CSV | Passed | 1807.8 ms | 2449 |
| JSON | Passed | 266.8 ms | 4133 |
| PDF | Passed | 389.4 ms | 4455 |
| DOCX | Passed | 347.6 ms | 2259 |

The report recorded `overallSuccess=true` and a **cumulative process** peak
of 98,078,720 bytes (~93.5 MiB). Those timings are not independent-process
benchmarks; CSV includes possible startup/first-conversion costs.

The probe's default worksheet selection was used: this test does **not**
verify every worksheet, data value, formula, visual element, or PDF parity
with Excel. The original file stays off GitHub. A validated production
deployment still needs application-specific limits and manual API review.

**Runtime lifecycle:** The package may retain `net8.0` compatibility, but
the .NET 8 runtime reached Microsoft's support end on 2026-11-10. Prefer
`net10.0` and a supported runtime for new production deployments.

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
concurrency, or hardware. The UK ONS real-world sample has now passed local, first-sheet structural
conversion checks. This does not establish comprehensive fidelity. Independent
synthetic Excel PDF comparisons have also been reviewed and documented.

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
- Four original synthetic fixtures and a separate advanced print-settings fixture
  have been reviewed against Excel LTSC. The independently generated reference
  PDFs are not committed publicly, and layout differences remain.
- The 1K, 10K and 100K isolated benchmarks have been run; their observed
  medians are environment-specific, not performance guarantees.
