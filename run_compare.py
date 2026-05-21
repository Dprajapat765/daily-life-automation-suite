"""CLI entry point for Tool 01 — Excel Compare."""
from __future__ import annotations

import argparse
import sys

from tools.excel_compare import compare_files


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compare two Excel/CSV files and produce a diff report."
    )
    parser.add_argument("file1", help="Baseline file (.xlsx or .csv)")
    parser.add_argument("file2", help="Comparison file (.xlsx or .csv)")
    parser.add_argument(
        "--key", nargs="+", metavar="COL", help="Key column(s) for row matching"
    )
    parser.add_argument(
        "--output", default="outputs/comparison_report.xlsx", help="Output report path"
    )
    parser.add_argument(
        "--no-auto-key",
        action="store_true",
        help="Disable automatic key column detection",
    )
    args = parser.parse_args()

    try:
        result = compare_files(
            args.file1,
            args.file2,
            key_columns=args.key,
            auto_key=not args.no_auto_key,
        )
        out = result.export_report(args.output)
        print(f"Report written to: {out}")
        print(f"Summary: {result.summary}")
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
