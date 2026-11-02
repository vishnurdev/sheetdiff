"""Command-line interface: `sheetdiff old.xlsx new.xlsx --key id`."""

from __future__ import annotations

import argparse
import sys

from . import __version__
from .core import SheetDiffError, compare_files
from .report import format_console, write_excel, write_html


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="sheetdiff",
        description="Compare two Excel/CSV files and report added, removed and changed rows.",
    )
    p.add_argument("left", help="Old / expected file (.csv, .tsv, .xlsx, .xls)")
    p.add_argument("right", help="New / actual file")
    p.add_argument("-k", "--key", action="append", required=True,
                   help="Key column that identifies a row (repeat for multi-column keys)")
    p.add_argument("--left-sheet", help="Sheet name in the left Excel file (default: first)")
    p.add_argument("--right-sheet", help="Sheet name in the right Excel file (default: first)")
    p.add_argument("-t", "--tolerance", type=float, default=0.0,
                   help="Ignore numeric differences up to this amount (e.g. 0.01)")
    p.add_argument("-i", "--ignore-case", action="store_true", help="Compare text case-insensitively")
    p.add_argument("-x", "--ignore-column", action="append", default=[],
                   help="Column to skip (repeatable)")
    p.add_argument("--excel", metavar="PATH", help="Write an Excel report to PATH")
    p.add_argument("--html", metavar="PATH", help="Write an HTML report to PATH")
    p.add_argument("--fail-on-diff", action="store_true",
                   help="Exit with code 1 if any difference is found (useful in CI)")
    p.add_argument("-q", "--quiet", action="store_true", help="Don't print the console report")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = compare_files(
            args.left,
            args.right,
            args.key,
            left_sheet=args.left_sheet,
            right_sheet=args.right_sheet,
            tolerance=args.tolerance,
            ignore_case=args.ignore_case,
            ignore_columns=args.ignore_column,
        )
    except SheetDiffError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    if not args.quiet:
        print(format_console(result))
    if args.excel:
        print(f"Excel report saved: {write_excel(result, args.excel)}")
    if args.html:
        print(f"HTML report saved:  {write_html(result, args.html)}")

    return 1 if (args.fail_on_diff and not result.is_identical) else 0


if __name__ == "__main__":
    sys.exit(main())
