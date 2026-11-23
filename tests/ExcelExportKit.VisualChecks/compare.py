#!/usr/bin/env python3
"""Compare PDFs by rendered pixels, preserving actionable visual artifacts.

This is a reviewer-assist tool, not a claim of strict print fidelity.
Rendering uses Poppler's pdftoppm at fixed DPI and Pillow for pixel comparison.
The comparison report marks changed pages without failing intentional design changes.
"""
import argparse
import json
import pathlib
import subprocess
from PIL import Image, ImageChops

FIXTURES = ("merged-header", "vertical-merge", "multi-page", "image-anchor")


def render_folder(source: pathlib.Path, destination: pathlib.Path, dpi: int, required: bool = True) -> dict:
    destination.mkdir(parents=True, exist_ok=True)
    summaries = {}
    for fixture in FIXTURES:
        pdf = source / (fixture + ".pdf")
        if not pdf.is_file():
            if not required:
                continue  # Newly introduced fixtures are absent from older baselines.
            raise FileNotFoundError(pdf)
        prefix = destination / fixture
        subprocess.run(
            ["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(prefix)],
            check=True, stdout=subprocess.DEVNULL,
        )
        pages = sorted(destination.glob(fixture + "-*.png"))
        if not pages:
            raise RuntimeError(f"No pages rendered for {fixture}")
        if fixture == "multi-page" and len(pages) < 2:
            raise RuntimeError("Expected multi-page fixture to produce at least two PDF pages")

        # Text checking supplements, not replaces, visual inspection.
        result = subprocess.run(
            ["pdftotext", "-layout", str(pdf), "-"],
            check=True, text=True, capture_output=True,
        )
        expected_text = {
            "merged-header": ["Q4 Sales Overview", "Tashkent", "Samarkand"],
            "vertical-merge": ["Foundation", "Amina", "Laylo"],
            "multi-page": ["Student 1", "Student 119"],
            "image-anchor": ["Student", "Amina", "Ali"],
        }[fixture]
        missing = [item for item in expected_text if item not in result.stdout]
        if missing:
            raise RuntimeError(f"{fixture}: PDF text missing {missing}")
        summaries[fixture] = {
            "pages": len(pages), "sizeBytes": pdf.stat().st_size,
            "requiredTextPresent": True,
        }
    return summaries


def compare(
    before: pathlib.Path, after: pathlib.Path, diffs: pathlib.Path
) -> list:
    diffs.mkdir(parents=True, exist_ok=True)
    results = []
    for fixture in FIXTURES:
        left = sorted(before.glob(fixture + "-*.png"))
        right = sorted(after.glob(fixture + "-*.png"))
        page_count = max(len(left), len(right))
        for index in range(page_count):
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
                mask = delta.convert("L").point(lambda p: 255 if p else 0)
                hist = mask.histogram()
                changed = sum(hist[1:])
                fraction = round(changed / (a.width * a.height), 6)
                result.update(status="changed" if changed else "identical",
                              changedFraction=fraction,
                              size=[a.width, a.height])
                if changed:
                    # Black: unchanged, white: different pixel
                    mask.save(diffs / f"{fixture}-{index + 1:02d}-diff.png")
                results.append(result)
    return results


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidate", required=True, type=pathlib.Path)
    ap.add_argument("--baseline", type=pathlib.Path)
    ap.add_argument("--output", required=True, type=pathlib.Path)
    ap.add_argument("--dpi", type=int, default=110)
    args = ap.parse_args()
    if args.dpi < 50 or args.dpi > 200:
        ap.error("--dpi must be 50..200")
    args.output.mkdir(parents=True, exist_ok=True)

    candidate = render_folder(args.candidate, args.output / "candidate", args.dpi)
    results = {
        "dpi": args.dpi,
        "candidate": candidate,
        "baselineAvailable": bool(args.baseline),
    }
    if args.baseline:
        baseline = render_folder(args.baseline, args.output / "baseline", args.dpi, required=False)
        results["baseline"] = baseline
        results["comparison"] = compare(
            args.output / "baseline",
            args.output / "candidate",
            args.output / "diffs",
        )

    report = args.output / "visual-report.json"
    report.write_text(json.dumps(results, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(results, indent=2))


if __name__ == "__main__":
    main()
