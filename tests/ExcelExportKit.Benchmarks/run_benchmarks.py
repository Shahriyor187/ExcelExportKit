#!/usr/bin/env python3
"""Repeat conversions in separate warmed-up .NET worker processes.

The generated XLSX is kept on disk; each format/trial gets a fresh process.
Peak RSS includes that worker's warmup, not fixture generation or other formats.
"""
import argparse
import json
import pathlib
import statistics
import subprocess
import sys

def worker(dll: pathlib.Path, *args: str) -> dict:
    completed = subprocess.run(
        ["dotnet", str(dll), *args], check=True, capture_output=True, text=True
    )
    return json.loads(completed.stdout.strip())

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dll", required=True, type=pathlib.Path)
    parser.add_argument("--rows", required=True, type=int, choices=[1000, 10000, 100000])
    parser.add_argument("--output", required=True, type=pathlib.Path)
    parser.add_argument("--repeats", type=int, default=3)
    args = parser.parse_args()
    if not 1 <= args.repeats <= 10:
        parser.error("repeats must be 1..10")
    directory = args.output.resolve()
    dll = args.dll.resolve()
    directory.mkdir(parents=True, exist_ok=True)
    fixture = directory / "fixture.xlsx"
    generated = worker(dll, "generate", str(args.rows), str(fixture))
    formats = ["csv", "json"] if args.rows > 10000 else [
        "csv", "json", "pdf", "docx"
    ]

    result = {
        "rows": args.rows, "columns": 6, "repeats": args.repeats,
        "fixtureBytes": generated["bytes"],
        "methodology": (
            "Each format/trial uses a separate .NET 10 process. Each worker "
            "does one unmeasured warmup conversion, forces GC, then times one "
            "conversion. Reported peak working set includes worker startup and "
            "warmup, but excludes fixture generation and other formats. "
            "Measurements depend on runner resources and input characteristics."
        ),
        "results": {}
    }
    for fmt in formats:
        observations = []
        for trial in range(args.repeats):
            output = directory / f"result-{fmt}-{trial + 1}.{fmt}"
            sample = worker(dll, "measure", fmt, str(fixture), str(output))
            observations.append(sample)
            print(f"{fmt} #{trial + 1}: {sample['elapsedMs']:.1f} ms; "
                  f"peak RSS {sample['peakWorkingSetBytes'] / 1048576:.1f} MiB",
                  file=sys.stderr)
            output.unlink()
        result["results"][fmt] = {
            "samples": observations,
            "medianElapsedMs": round(statistics.median(
                x["elapsedMs"] for x in observations), 2),
            "medianPeakWorkingSetBytes": round(statistics.median(
                x["peakWorkingSetBytes"] for x in observations)),
            "medianAllocatedBytes": round(statistics.median(
                x["allocatedBytes"] for x in observations))
        }
        report = directory / "benchmark.json"
        report.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2))

if __name__ == "__main__":
    main()
