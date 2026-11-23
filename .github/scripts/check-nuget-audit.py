#!/usr/bin/env python3
"""Fail if a .NET package list JSON report contains any known advisories.

Run with:
  dotnet list ExcelExportKit.csproj package --vulnerable --include-transitive \
    --format json --output-version 1 --no-restore > audit.json
  python3 .github/scripts/check-nuget-audit.py audit.json

This is an advisory-feed audit, not a comprehensive security review.
"""
import json
import pathlib
import sys

if len(sys.argv) != 2:
    raise SystemExit("Usage: check-nuget-audit.py <dotnet-list-report.json>")
path = pathlib.Path(sys.argv[1])
data = json.loads(path.read_text(encoding="utf-8-sig"))
if not isinstance(data, dict) or not isinstance(data.get("projects"), list) or not data["projects"]:
    raise SystemExit("Audit report missing project data; fail closed.")

findings = []
framework_count = 0
for project in data["projects"]:
    for framework in project.get("frameworks", []):
        framework_count += 1
        for category in ("topLevelPackages", "transitivePackages"):
            for package in framework.get(category, []):
                for advisory in package.get("vulnerabilities", []):
                    findings.append({
                        "project": project.get("path"),
                        "framework": framework.get("framework"),
                        "package": package.get("id"),
                        "resolved": package.get("resolvedVersion"),
                        "severity": advisory.get("severity"),
                        "advisory": advisory.get("advisoryurl", advisory.get("advisoryUrl")),
                    })

if findings:
    print(json.dumps(findings, indent=2))
    raise SystemExit(f"Release blocked: {len(findings)} reported NuGet vulnerabilities.")

print(f"NuGet advisory check: no reported vulnerabilities across "
      f"{len(data['projects'])} project(s), {framework_count} framework entries.")
