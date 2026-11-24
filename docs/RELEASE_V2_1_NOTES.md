# ExcelExportKit v2.1.0 — draft release notes

Not published. Final release depends on Phase 1.8 verification.

## Added
- Dedicated TSV conversion, including batch and all-sheet operations.
- Stream-to-stream, path-to-stream and stream-to-path overloads for CSV, TSV, JSON, PDF and DOCX.
- Explicit non-seekable input buffering modes with bounded memory or temporary-file buffering.
- RowOffset, MaxRows and configurable progress reporting.

## Improved
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

See README and docs/PRODUCTION_GUIDE.md. Final test counts, benchmark observations and dependency audit results will be recorded after release validation.
