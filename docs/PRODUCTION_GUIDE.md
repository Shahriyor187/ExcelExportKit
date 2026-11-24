# Production integration guide (v2.1 candidate)

These examples require the v2.1 source or the corresponding released package. Prefer a supported .NET 10 runtime.

## Authenticated download from a trusted workbook

This example uses an application-owned workbook path, not a user-supplied path. Configure authentication and the Reports.Read authorization policy before registering the endpoint.

```csharp
using ExportFiles.Services;

app.MapGet("/reports/sales.csv", async (HttpContext context) =>
{
    context.Response.ContentType = "text/csv; charset=utf-8";
    context.Response.Headers.ContentDisposition = "attachment; filename=sales.csv";
    context.Response.Headers.CacheControl = "no-store";

    await ExcelConverter.ToCSVAsync("reports/sales.xlsx",
        context.Response.Body, new ExcelConversionOptions
        {
            CsvProtectFormulas = true,
            MaxWorkbookFileBytes = 20 * 1024 * 1024,
            MaxUncompressedWorkbookBytes = 80 * 1024 * 1024,
            MaxWorkbookArchiveEntries = 1024,
            MaxSelectedCells = 100_000
        }, context.RequestAborted);
}).RequireAuthorization("Reports.Read");
```

Direct HTTP output cannot be rolled back. If conversion fails after headers are sent, abort the response instead of appending a JSON error to the file. For an all-or-nothing download, finish conversion to a private staged file before starting the response. The application must enforce access control and tenant isolation.

## Handling uploaded workbooks

Treat uploads as untrusted even when their extension, MIME type or ZIP signature looks correct. Enforce a request-body limit before multipart parsing. Authenticate, authorize and rate-limit uploads; cap concurrent conversions, queue length and disk usage. Run hostile workbook parsing and rendering in a separate least-privilege worker with OS-level memory, CPU, disk and wall-clock limits, no unnecessary network access, and no application secrets. Do not rely on cancellation or ZIP metadata as a sandbox.

Use an application-controlled private temporary directory. On Unix restrict the parent directory; on Windows configure restrictive ACLs. Never use the submitted filename as a filesystem path. Keep input and output separate and do not modify them concurrently. The following is a worker-side conversion, not a complete upload endpoint:

```csharp
using ExportFiles.Services;

static Task ConvertUploadAsync(Stream upload, Stream output,
    string privateDirectory, CancellationToken ct) =>
    ExcelConverter.ToCSVAsync(upload, output,
        new ExcelConversionOptions
        {
            InputBufferingMode = InputBufferingMode.TemporaryFile,
            TemporaryDirectory = privateDirectory,
            MaxBufferedInputBytes = 10 * 1024 * 1024,
            MaxWorkbookFileBytes = 10 * 1024 * 1024,
            MaxUncompressedWorkbookBytes = 40 * 1024 * 1024,
            MaxWorkbookArchiveEntries = 512,
            MaxSelectedCells = 100_000,
            CsvProtectFormulas = true
        }, ct);
```

Seekable inputs must start at zero. Seekable output must be empty and at zero. Non-seekable input is rejected by default; explicit buffering is bounded during copying. Stream API defaults are 64 MiB compressed input, 256 MiB expanded archive, 4096 entries and 1,000,000 selected cells. Legacy path APIs retain optional limits. Set tighter limits for your workload. Limit the size of generated files separately, especially PDF and DOCX; input limits do not bound output size or process memory.

ClosedXML loads workbooks into memory; stream support is not a forward-only XLSX parser. PDF and Word use private staging, which consumes disk space. Embedded images are opt-in and restricted by count, byte size and original/displayed dimensions; this does not guarantee safe decoding of hostile files. Disable them when not needed. Print sections have bounded counts and an optional cumulative cell budget. Keep dependency versions current and review advisory feeds.

## Cancellation and cleanup

Caller streams remain open. Discard partial output on cancellation or failure. The parser and renderer contain synchronous boundaries. Temporary files are normally removed on success, failure or cooperative cancellation, but an abrupt process termination can leave artifacts. Use private storage, quotas and an external cleanup policy. Progress reports measure processed rows rather than successful publication. Asynchronous Progress<T> callbacks run independently. Avoid logging workbook contents or sensitive paths.

## Reproducing progress overhead

```bash
dotnet run -c Release --project tests/ExcelExportKit.Benchmarks -- --progress
```

For PDF fidelity and resource measurements, use the separate visual regression and benchmark workflows. They do not establish exact Microsoft Excel print fidelity or a guaranteed security boundary.
