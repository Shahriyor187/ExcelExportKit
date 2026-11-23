# Independent Microsoft Excel PDF reference review — 2026-11-23

**Scope:** synthetic visual regression fixtures exported independently from
Microsoft Excel LTSC on Windows, compared with the ExcelExportKit PDF workflow
from the `2.0.0-rc.1` branch/commit on `master` (run [45149410131](https://github.com/Shahriyor187/ExcelExportKit/actions/runs/45149410131)).
The independently supplied PDFs report `Microsoft® Excel® LTSC` as the PDF
producer/creator. This review examines content and visible layout, **not exact
binary or pixel matching**.

The original four workbook fixtures are in the visual regression job artifacts.
The independently exported Excel PDFs were reviewed externally and were **not
committed into the public repository**.

## Verification results

| Fixture | ExcelExportKit | Excel LTSC | Content outcome |
| --- | --- | --- | --- |
| `merged-header` | 1 A4 landscape page | 1 A4 portrait page | All headings, regions and values present |
| `vertical-merge` | 1 portrait page | 1 portrait page | Four students, scores and merged labels present |
| `multi-page` | 5 portrait pages | 3 portrait pages | **119/119** student, score and status records match |
| `image-anchor` | 1 portrait page, 2 images | 1 portrait page, 2 images | Both PNG and JPEG visibly present |

### Print pagination

- ExcelExportKit page data distribution: **29 / 29 / 29 / 29 / 3** rows.
- Excel LTSC page distribution: **47 / 48 / 24** rows.
- The library repeats the table header on all five pages. The workbook's
  standard Excel export prints the table heading on its first page only.
- The library's fixture generator deliberately sets `PdfLandscape=true`
  for `merged-header`, while the Excel workbook itself prints portrait.
  This is an option mismatch, not evidence that landscape is malfunctioning.
- The library also intentionally adds `DocumentTitle = "Visual fixture: ..."`
  and page-number footers not present in the native Excel exports.

### Other layout differences

- Text size, row spacing, horizontal scaling and column proportions differ.
- Vertically merged label placement differs. Data and outside borders remain.
- Cell-anchored pictures are placed *inside* PDF table cells by the library;
  Excel preserves drawing anchors, with visibly different placement and
  neighboring text flow. This limitation is already described in README.
- Alternating row fills and basic header colors are present in both reports.

## Interpretation and release policy

**Content-level fidelity on these four fixtures: PASS.**
Expected labels and all 119 multi-page records match; neither embedded
picture is missing. This is a narrowly scoped synthetic test, not broad
file-format conformance.

**Pixel-perfect Microsoft Excel print fidelity: NOT ESTABLISHED.**
Print pagination and layout differ visibly. Do not advertise exact Excel
export compatibility or claim layout identity based on the existing
visual-regression workflow.

The `2.0.0` release can retain its documented **approximate** PDF layout
contract. The following gates still need explicit review before stable release:

1. One or more representative *real-world* workbooks (with authorization and
   no user-sensitive files uploaded to public CI artifacts).
2. Separate Excel LTSC references for the **opt-in worksheet print-layout**
   mode: print areas, scale, margins, repeated print titles, both kinds of
   manual page break and landscape.
3. Human visual inspection of images/pages, ideally across Excel versions and
   font environments, with known limitations explicitly retained in the README.

Prioritize narrower consistency improvements rather than changing default
pagination: matching Excel's 3-page output could require different row-height
and table-scale behavior that would change previous PDF output. Any changes
must have a clear opt-in compatibility strategy and new reference tests.

**Conclusion:** the independent Excel-reference step is partially complete
(merged cells, 119-row report and pictures). The separate advanced
print-settings reference test remains outstanding.
