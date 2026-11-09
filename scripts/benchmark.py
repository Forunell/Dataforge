"""Repeatable synthetic local benchmark. No real files or network calls."""
import csv
import io
import json
import platform
import time
from pathlib import Path

from openpyxl import Workbook

from dataforge.core import analyze, export_files, read_source, suggest

ROOT = Path(__file__).resolve().parents[1]


def main():
    cases = []
    for label, columns, long_names, xlsx in [("csv-2000", 5, False, False), ("xlsx-2000", 5, False, True),
                                            ("csv-2000x64", 64, False, False), ("csv-near-2mb", 5, True, False)]:
        header = ["sku", "name", "price", "currency", "stock"] + [f"extra{i}" for i in range(columns - 5)]
        rows = [[f"{i:06}", "Tea " + ("x" * 900 if long_names else str(i)), "1234.56", "UAH", "10"] + ["text"] * (columns - 5) for i in range(2000)]
        if xlsx:
            book = Workbook()
            book.active.append(header)
            for row in rows:
                book.active.append(row)
            stream = io.BytesIO()
            book.save(stream)
            payload = stream.getvalue()
        else:
            stream = io.StringIO(newline="")
            writer = csv.writer(stream, delimiter=";")
            writer.writerow(header)
            writer.writerows(rows)
            payload = stream.getvalue().encode("utf-8-sig")
        started = time.perf_counter()
        suffix = ".xlsx" if xlsx else ".csv"
        sources = [read_source("A" + suffix, payload), read_source("B" + suffix, payload)]
        parsed = time.perf_counter()
        result = analyze(sources, [suggest(s) for s in sources], {}, True)
        checked = time.perf_counter()
        files = export_files(result)
        finished = time.perf_counter()
        assert len(result["products"]) == 4000 and result["comparison_counts"] == {"unchanged": 2000}
        assert result["errors"] == 0 and "normalized.csv" in files
        cases.append(dict(case=label, rows_per_file=2000, columns=columns, bytes_per_file=len(payload),
                          parse_seconds=round(parsed-started, 3), analyze_seconds=round(checked-parsed, 3),
                          export_seconds=round(finished-checked, 3), total_seconds=round(finished-started, 3),
                          zip_bytes=len(files["dataforge-results.zip"]), status="PASS"))
    report = dict(date="2026-11-07", python=platform.python_version(), platform=platform.platform(), cases=cases,
                  scope="Single run per synthetic case on this machine. Not a universal speed or memory guarantee.")
    (ROOT / "artifacts/performance.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
