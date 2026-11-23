# Changelog

## 2.0.0 — Stable package prepared (2026-11-23; publication pending)

This entry describes the **stable package candidate in source control**.
It is **not** proof that `2.0.0` is available on NuGet.org. Publishing is
a separate, manually approved operation.

### New in v2
- Formatting-aware PDF output, fully selected horizontal/vertical merged cells,
  supported fonts/borders/colors, and opt-in cell-anchored PNG/JPEG images.
- Opt-in worksheet print areas, custom margins, scale percentage, repeating
  print title rows, and manual horizontal/vertical page breaks. Microsoft Excel
  pixel-perfect rendering, fit-to-page and charts remain unsupported.
- Worksheet selection, typed JSON, CSV formula protection, Word formatting
  options, and batch/multi-sheet export helpers.
- Opt-in workbook preflight limits for file bytes, expanded ZIP bytes, ZIP entry
  count, selected cell count, and image size/count. Callers remain responsible
  for timeouts, process isolation, concurrency and output quotas.
- Incremental CSV/JSON writing with reduced duplicate buffering; XLSX reading
  still uses in-memory ClosedXML.
- An isolated 1K/10K/100K benchmark harness, five visual PDF regression fixtures,
  independently reviewed Microsoft Excel LTSC PDF references, and a local-only
  real-workbook release probe with sanitized diagnostic reporting.
- End-to-end synthetic business workbook tests, plus clean packed-package
  consumer verification for .NET 8 and .NET 10 and automated dependency audits.

### Release evidence and limitations
- Four original visual fixtures and the separate advanced print-layout fixture
  were compared against Microsoft Excel LTSC. Data and manual section ordering
  match on those fixtures; pagination, widths and scaling visibly differ.
- An independently authored public UK ONS GDP workbook successfully converted
  to all four output types in the local probe. The probe validates output
  structure, not full per-cell equivalence or visual fidelity, and by default
  processes only the first worksheet.
- The `net8.0` target remains for consumer compatibility, but Microsoft ended
  .NET 8 runtime support on November 10, 2026. Prefer the `net10.0` runtime
  for supported production deployments.
- `2.0.0` is a **major version**; consumers should verify API semantics,
  especially opt-in options and their defaults, before upgrading.

### Release status
- Version and release notes are prepared for final CI/PDF and package inspection.
- NuGet publication, GitHub tagging and final release approval are **pending**.

## 1.2.0 development notes (incorporated into v2)

### Added
- Select worksheets by name or zero-based index.
- Optional strongly typed JSON values for numbers, dates, and Booleans.
- Real paginated PDF tables with repeated column headers, borders, and page numbers.
- PDF landscape orientation, Word document titles, table styles, and landscape pages.
- Convert every sheet or several workbooks with per-file success/failure results.
- Configurable CSV delimiters and formula-injection protection in options-based calls.
- ASP.NET Core PDF download example and expanded conversion tests.

### Compatibility
- Original two-argument conversion methods are retained.
- Legacy JSON remains string-oriented unless PreserveJsonTypes is enabled.
- Original CSV output behavior is preserved in the two-argument call.

## 1.1.0 — 2026-11-22

- Modernized packaging and dependency choices for .NET 8 and .NET 10.
- Added tests, docs, CI and NuGet Trusted Publishing.
