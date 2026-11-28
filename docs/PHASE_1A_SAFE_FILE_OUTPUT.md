# Phase 1A — Safe file publication (draft)

New opt-in API; legacy APIs and package versions remain unchanged.

```csharp
var result = await ExcelConverter.ConvertFileAsync("input.xlsx", "report.csv", "csv",
    maxOutputBytes: 10 * 1024 * 1024, cancellationToken: token);
Console.WriteLine(result.OutputBytes);
```

The destination directory must already exist and be application-controlled. The file is rendered in a private sibling staging directory, closed, checked against the optional positive final-size limit and published with a same-filesystem replacement rename. Cancellation is checked immediately before publication; cancellation racing with publication may still succeed. The result is returned only after successful publication. The optional limit measures finalized file bytes, including headers and container overhead, **not** a hard limit on temporary disk usage, allocations or memory. Use OS quotas and worker isolation for those. This is not a crash-durability or cross-filesystem transaction guarantee. Cleanup is best effort; arrange cleanup of abandoned stage directories after crashes or filesystem errors. On Windows, protection depends on destination directory ACLs. Do not use attacker-controlled paths.

The API uses the existing stream converters and their finite input safety defaults. PDF and DOCX can create their own intermediate render files. The input/output paths must be distinct; path equality checking is lexical, not a filesystem identity guarantee.

Next: hard byte budgets where compatible with writers, typed row diagnostics and validation, reusable column mapping and conversion policies. Verify all tests on .NET 8 and .NET 10 before release.
