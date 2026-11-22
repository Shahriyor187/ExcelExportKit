# Changelog

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
