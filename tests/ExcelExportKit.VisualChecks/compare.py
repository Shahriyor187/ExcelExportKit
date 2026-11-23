#!/usr/bin/env python3
"""Render PDF fixtures, validate visible content, and compare page pixels.

The master-branch baseline detects regressions in this renderer. Optional
Excel-exported reference PDFs provide independent (human-reviewed) context,
not an assertion of pixel-exact Excel equivalence.
"""
import argparse
import json
import pathlib
import subprocess
from PIL import Image, ImageChops

FIXTURES = ("merged-header", "vertical-merge", "multi-page", "image-anchor", "print-layout")
EXPECTED_TEXT = {
    "merged-header": ("Q4 Sales Overview", "Tashkent", "Samarkand"),
    "vertical-merge": ("Foundation", "Amina", "Laylo"),
    "multi-page": ("Student 1", "Student 119"),
    "image-anchor": ("Student", "Amina", "Ali"),
    "print-layout": ("Item 1", "Item 32", "Print section"),
}


def image_color_counts(path: pathlib.Path) -> dict:
    """Look for the two deliberately saturated, distinguishable fixture images."""
    magenta = green = 0
    with Image.open(path) as picture:
        for red, g, blue in picture.convert("RGB").getdata():
            if red >= 180 and g <= 110 and blue >= 130:
                magenta += 1
            if g >= 145 and red <= 115 and blue <= 115:
                green += 1
    return {"pngMagentaPixels": magenta, "jpegGreenPixels": green}


def verify_image_fixture(path: pathlib.Path) -> dict:
    counts = image_color_counts(path)
    # Each colored picture occupies thousands of pixels at 110 DPI. Requiring
    # only 50 tolerates rasterization/antialiasing without accepting missing
    # images or the earlier 1x1 white pixel fixture.
    if min(counts.values()) < 50:
        raise RuntimeError(
            "image-anchor: expected visible magenta PNG and green JPEG pixels; "
            f"observed {counts}"
        )
    return counts



def verify_print_layout_pages(pdf: pathlib.Path, page_count: int) -> None:
    """Guard the expected two row segments x two column segments page order.

    Independent Excel LTSC references confirmed this four-page section ordering.
    This does NOT imply identical Excel geometry, page margins or font scaling.
    """
    import re
    if page_count != 4:
        raise RuntimeError(f"print-layout: expected 4 manual sections, got {page_count}")
    expected = (
        ("Section", "Report title row", "Part A", "Item 1", "Item 16"),
        ("Section", "Report title row", "Part B", "Item 17", "Item 32"),
        ("Score", "Scale 85%", "Landscape", "1109", "1244"),
        ("Score", "Scale 85%", "Landscape", "1253", "1388"),
    )
    for index, required in enumerate(expected, start=1):
        page_text = subprocess.run(
            ["pdftotext", "-f", str(index), "-l", str(index),
             "-layout", str(pdf), "-"],
            check=True, capture_output=True, text=True,
        ).stdout
        # Whole-token matching avoids accepting "Item 1" as a substring of
        # "Item 10" when testing print sections.
        missing = [
            phrase for phrase in required
            if not re.search(r"(?<!\\w)" + re.escape(phrase) + r"(?!\\w)", page_text)
        ]
        if missing:
            raise RuntimeError(f"print-layout page {index}: missing {missing}")

def render_folder(
    source: pathlib.Path,
    destination: pathlib.Path,
    dpi: int,
    required: bool = True,
    check_text: bool = True,
    check_images: bool = False,
) -> dict:
    destination.mkdir(parents=True, exist_ok=True)
    summaries = {}
    for fixture in FIXTURES:
        pdf = source / (fixture + ".pdf")
        if not pdf.is_file():
            if not required:
                continue  # New fixtures may be absent from older baselines/references.
            raise FileNotFoundError(pdf)
        prefix = destination / fixture
        subprocess.run(
            ["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(prefix)],
            check=True, stdout=subprocess.DEVNULL,
        )
        pages = sorted(destination.glob(fixture + "-*.png"))
        if not pages:
            raise RuntimeError(f"No pages rendered for {fixture}")
        if check_text:
            if fixture == "multi-page" and len(pages) < 2:
                raise RuntimeError("Expected multi-page fixture to produce at least two PDF pages")
            result = subprocess.run(
                ["pdftotext", "-layout", str(pdf), "-"],
                check=True, text=True, capture_output=True,
            )
            missing = [text for text in EXPECTED_TEXT[fixture] if text not in result.stdout]
            if missing:
                raise RuntimeError(f"{fixture}: PDF text missing {missing}")
            if fixture == "print-layout":
                verify_print_layout_pages(pdf, len(pages))

        summary = {"pages": len(pages), "sizeBytes": pdf.stat().st_size}
        if check_text:
            summary["requiredTextPresent"] = True
        if check_images and fixture == "image-anchor":
            summary.update(verify_image_fixture(pages[0]))
        summaries[fixture] = summary
    return summaries


def compare(
    before: pathlib.Path,
    after: pathlib.Path,
    diffs: pathlib.Path,
    missing_before_status: str = "new-fixture",
) -> list:
    diffs.mkdir(parents=True, exist_ok=True)
    results = []
    for fixture in FIXTURES:
        left = sorted(before.glob(fixture + "-*.png"))
        right = sorted(after.glob(fixture + "-*.png"))
        if not left and right:
            results.append({
                "fixture": fixture, "page": None,
                "status": missing_before_status, "newPages": len(right)
            })
            continue
        for index in range(max(len(left), len(right))):
            result = {"fixture": fixture, "page": index + 1}
            if index >= len(left) or index >= len(right):
                result.update(status="page-count-changed", changedFraction=1.0)
                results.append(result)
                continue
            with Image.open(left[index]) as a, Image.open(right[index]) as b:
                a = a.convert("RGB")
                b = b.convert("RGB")
                if a.size != b.size:
                    result.update(status="dimension-changed", oldSize=a.size, newSize=b.size)
                    results.append(result)
                    continue
                delta = ImageChops.difference(a, b)
                # Use the maximum RGB channel difference; converting to L first
                # can round away subtle single-channel changes.
                red, green, blue = delta.split()
                mask = ImageChops.lighter(ImageChops.lighter(red, green), blue)
                mask = mask.point(lambda p: 255 if p else 0)
                hist = mask.histogram()
                changed = sum(hist[1:])
                fraction = round(changed / (a.width * a.height), 6)
                result.update(
                    status="changed" if changed else "identical",
                    changedFraction=fraction, size=[a.width, a.height]
                )
                if changed:
                    mask.save(diffs / f"{fixture}-{index + 1:02d}-diff.png")
                results.append(result)
    return results


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidate", required=True, type=pathlib.Path)
    ap.add_argument("--baseline", type=pathlib.Path)
    ap.add_argument("--excel-reference", type=pathlib.Path,
                    help="Optional directory of PDFs exported independently by Microsoft Excel")
    ap.add_argument("--output", required=True, type=pathlib.Path)
    ap.add_argument("--dpi", type=int, default=110)
    args = ap.parse_args()
    if args.dpi < 50 or args.dpi > 200:
        ap.error("--dpi must be 50..200")
    args.output.mkdir(parents=True, exist_ok=True)

    candidate = render_folder(
        args.candidate, args.output / "candidate", args.dpi, check_images=True
    )
    results = {
        "dpi": args.dpi,
        "candidate": candidate,
        "baselineAvailable": bool(args.baseline),
    }
    if args.baseline:
        baseline = render_folder(
            args.baseline, args.output / "baseline", args.dpi, required=False
        )
        results["baseline"] = baseline
        results["comparison"] = compare(
            args.output / "baseline",
            args.output / "candidate",
            args.output / "diffs",
        )

    if args.excel_reference:
        reference = render_folder(
            args.excel_reference,
            args.output / "excel-reference",
            args.dpi,
            required=False,
            check_text=False,  # Excel's own printing and pagination can differ.
        )
        if not reference:
            raise FileNotFoundError(
                f"No recognized PDF fixtures in {args.excel_reference}"
            )
        results["excelReference"] = reference
        results["excelComparison"] = compare(
            args.output / "excel-reference",
            args.output / "candidate",
            args.output / "excel-diffs",
            missing_before_status="missing-excel-reference",
        )

    report = args.output / "visual-report.json"
    report.write_text(json.dumps(results, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(results, indent=2))


if __name__ == "__main__":
    main()
