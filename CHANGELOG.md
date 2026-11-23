# Changelog

## 2.0.0-rc.1 — Release candidate (not published; 2026-11-23)

This is a **candidate build for testing**, not the final stable release.

### Development changes
- Styled PDF conversion with merged cells, border/font styles, images and an opt-in print-layout mode.
- PDF visual regression checks with optional independently generated Excel PDF references.
- Opt-in XLSX input safeguards: compressed and expanded file sizes, archive entry count and selected worksheet cell count.
- Fix image-only styled PDF rows being accidentally omitted by `SkipEmptyRows`.
- Reproducible 1K/10K/100K synthetic conversion benchmark harness and manual GitHub Actions benchmark workflow.
- Realistic invoice and larger report integration fixtures.
- Release-readiness checklist for performance/security review and v2 packaging.
- Incremental CSV and JSON output writing, reduced duplicate typed-cell allocations, and isolated per-format benchmarks.
- Real packed-package .NET 8 and .NET 10 consumer tests, direct/transitive NuGet advisory checks, and guarded publishing workflow.

### Status
- Project version is now `2.0.0-rc.1` on the candidate branch only; no NuGet publication is performed by this PR.
- Independent Excel reference PDFs, representative user-workbook testing and release approval remain outstanding gates.
- Before stable publication, change version to `2.0.0` in a separate approved release commit.

## 1.2.0 — Unreleased

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
