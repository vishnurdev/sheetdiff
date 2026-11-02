"""SheetDiff - compare and validate Excel/CSV files."""

__version__ = "0.1.0"

from .core import DiffResult, SheetDiffError, compare, compare_files, load_table  # noqa: E402
from .report import format_console, write_excel, write_html  # noqa: E402

__all__ = [
    "DiffResult", "SheetDiffError", "compare", "compare_files", "load_table",
    "format_console", "write_excel", "write_html",
]
