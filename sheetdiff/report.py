"""Excel, HTML and console reports for a DiffResult."""

from __future__ import annotations

import html
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .core import DiffResult

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF")
ADDED_FILL = PatternFill("solid", fgColor="DCFCE7")
REMOVED_FILL = PatternFill("solid", fgColor="FEE2E2")
CHANGED_FILL = PatternFill("solid", fgColor="FEF3C7")


# --------------------------------------------------------------------------- #
# Console
# --------------------------------------------------------------------------- #
def format_console(result: DiffResult, max_changes: int = 20) -> str:
    lines = ["", "SheetDiff report", "=" * 40]
    width = max(len(k) for k in result.summary())
    for k, v in result.summary().items():
        lines.append(f"{k.ljust(width)} : {v}")

    if not result.totals.empty:
        lines += ["", "Column totals", "-" * 40, result.totals.to_string(index=False)]

    if not result.changed.empty:
        lines += ["", f"Changed cells (first {max_changes})", "-" * 40]
        lines.append(result.changed.head(max_changes).to_string(index=False))
    if not result.added.empty:
        lines += ["", "Added rows", "-" * 40, result.added.head(max_changes).to_string(index=False)]
    if not result.removed.empty:
        lines += ["", "Removed rows", "-" * 40, result.removed.head(max_changes).to_string(index=False)]
    return "\n".join(lines) + "\n"


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
def _style_sheet(ws, row_fill: PatternFill | None = None) -> None:
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    if row_fill is not None:
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.fill = row_fill
    for i, col in enumerate(ws.columns, start=1):
        longest = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(i)].width = min(max(longest + 2, 10), 60)
    ws.freeze_panes = "A2"


def write_excel(result: DiffResult, path: str | Path) -> Path:
    """Write a multi-sheet Excel report: Summary, Totals, Changed, Added, Removed."""
    path = Path(path)
    summary = pd.DataFrame(list(result.summary().items()), columns=["Metric", "Value"])
    sheets = [
        ("Summary", summary, None),
        ("Totals", result.totals, None),
        ("Changed", result.changed, CHANGED_FILL),
        ("Added", result.added, ADDED_FILL),
        ("Removed", result.removed, REMOVED_FILL),
    ]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df, _ in sheets:
            out = df if not df.empty else pd.DataFrame({"Info": ["No rows"]})
            out.to_excel(writer, sheet_name=name, index=False)
        for name, df, fill in sheets:
            _style_sheet(writer.sheets[name], fill if not df.empty else None)
    return path


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #
_CSS = """
:root{--bg:#f8fafc;--card:#fff;--text:#0f172a;--muted:#64748b;--line:#e2e8f0;
--add:#dcfce7;--rem:#fee2e2;--chg:#fef3c7;--accent:#2563eb}
@media (prefers-color-scheme:dark){:root{--bg:#0b1120;--card:#111827;--text:#e5e7eb;
--muted:#94a3b8;--line:#1f2937;--add:#14532d;--rem:#7f1d1d;--chg:#713f12;--accent:#60a5fa}}
*{box-sizing:border-box}body{margin:0;font:14px/1.5 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;
background:var(--bg);color:var(--text)}main{max-width:1100px;margin:0 auto;padding:32px 16px}
h1{margin:0 0 4px;font-size:26px}h2{margin:32px 0 10px;font-size:18px}
.sub{color:var(--muted);margin:0 0 24px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px}
.card b{display:block;font-size:24px}.card span{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.04em}
.badge{display:inline-block;padding:4px 10px;border-radius:999px;font-weight:600;font-size:12px}
.ok{background:var(--add)}.bad{background:var(--chg)}
.wrap{overflow-x:auto;background:var(--card);border:1px solid var(--line);border-radius:12px}
table{border-collapse:collapse;width:100%}th,td{padding:8px 12px;border-bottom:1px solid var(--line);
text-align:left;white-space:nowrap}th{font-size:12px;text-transform:uppercase;color:var(--muted)}
tr:last-child td{border-bottom:0}.added td{background:var(--add)}.removed td{background:var(--rem)}
.changed td{background:var(--chg)}.empty{padding:14px;color:var(--muted)}
footer{margin-top:40px;color:var(--muted);font-size:12px}
"""


def _fmt(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float):
        return f"{value:,.4f}".rstrip("0").rstrip(".")
    return str(value)


def _table(df: pd.DataFrame, row_class: str = "") -> str:
    if df.empty:
        return '<div class="wrap"><div class="empty">No rows</div></div>'
    head = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
    rows = []
    for _, r in df.iterrows():
        cells = "".join(f"<td>{html.escape(_fmt(v))}</td>" for v in r)
        rows.append(f'<tr class="{row_class}">{cells}</tr>')
    return f'<div class="wrap"><table><thead><tr>{head}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>'


def write_html(result: DiffResult, path: str | Path) -> Path:
    """Write a standalone, shareable HTML report."""
    path = Path(path)
    status = (
        '<span class="badge ok">Identical</span>'
        if result.is_identical
        else '<span class="badge bad">Differences found</span>'
    )
    cards = [
        ("Rows in left", result.left_rows),
        ("Rows in right", result.right_rows),
        ("Added", len(result.added)),
        ("Removed", len(result.removed)),
        ("Rows changed", result.changed_rows),
        ("Cells changed", len(result.changed)),
    ]
    cards_html = "".join(f'<div class="card"><span>{k}</span><b>{v}</b></div>' for k, v in cards)
    col_note = ""
    if result.columns_only_left or result.columns_only_right:
        col_note = (
            f"<p>Columns only in left: <b>{html.escape(', '.join(result.columns_only_left) or '-')}</b> · "
            f"only in right: <b>{html.escape(', '.join(result.columns_only_right) or '-')}</b></p>"
        )

    doc = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SheetDiff Report</title><style>{_CSS}</style></head><body><main>
<h1>SheetDiff Report {status}</h1>
<p class="sub">{html.escape(result.left_name)} → {html.escape(result.right_name)} ·
key: {html.escape(', '.join(result.key))}</p>
<div class="cards">{cards_html}</div>{col_note}
<h2>Column totals</h2>{_table(result.totals)}
<h2>Changed cells</h2>{_table(result.changed, "changed")}
<h2>Added rows</h2>{_table(result.added, "added")}
<h2>Removed rows</h2>{_table(result.removed, "removed")}
<footer>Generated by SheetDiff on {datetime.now():%Y-%m-%d %H:%M}</footer>
</main></body></html>"""
    path.write_text(doc, encoding="utf-8")
    return path
