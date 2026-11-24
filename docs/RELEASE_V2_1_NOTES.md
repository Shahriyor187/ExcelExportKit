# ExcelExportKit v2.1.0 — release notes

Stable release candidate; publication is pending final validation and explicit approval.

## Added
- Dedicated TSV conversion, including batch and all-sheet operations.
- Stream-to-stream, path-to-stream and stream-to-path overloads for CSV, TSV, JSON, PDF and DOCX.
- Explicit non-seekable input buffering modes with bounded memory or temporary-file buffering.
- RowOffset, MaxRows and configurable progress reporting.

## Improved
- Expanded archive validation limits reads to the remaining byte budget plus a single overflow-detection byte.
- Cancellation reaches individual conversions in batch operations.
- Finite default workbook limits for new stream APIs; shared archive validation and private staging.
- Staged publication for cancellation-aware and progress-enabled path conversions.
- Print-section resource accounting and original embedded-image dimension checks.
- Formula mitigation whitespace inspection avoids temporary trimmed strings.
- Production integration guidance and expanded benchmark verification.

## Compatibility and security
- Targets net8.0 and net10.0; prefer a supported runtime for production.
- Legacy path defaults remain compatible; new stream APIs apply finite limits by default.
- Caller streams remain open. Failed stream outputs must be discarded.
- ClosedXML remains an in-memory parser. Input limits do not guarantee process memory or output size bounds; isolate hostile uploads and enforce OS quotas.
- PDF and DOCX stream exports require temporary staging; Windows privacy depends on parent ACLs.
- Progress does not guarantee finalization; cancellation cannot interrupt every synchronous parser or renderer call.
- Exact Excel print fidelity is not claimed.

## Release candidate validation (November 24, 2026)

- 131 regression cases passed on each of .NET 8 and .NET 10 (262 executions).
- PDF visual regression and package consumer tests passed.
- No known vulnerabilities reported across 18 resolved dependency entries.
- Package validation against published v2.0.0 passed.
- Refreshed cross-package integration passed on .NET 8 and .NET 10. The final resource-boundary mutation was detected on both frameworks.
- Stable-version package validation remains a release gate. No package has been published.

See README and docs/PRODUCTION_GUIDE.md.
