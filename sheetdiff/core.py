"""Core comparison logic for SheetDiff."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd

SUPPORTED_EXTENSIONS = {".csv", ".tsv", ".txt", ".xlsx", ".xlsm", ".xls"}


class SheetDiffError(Exception):
    """Raised when inputs can't be compared (missing file, bad key, etc.)."""


@dataclass
class DiffResult:
    """Everything SheetDiff found when comparing two tables."""

    left_name: str
    right_name: str
    key: list[str]
    left_rows: int
    right_rows: int
    added: pd.DataFrame
    removed: pd.DataFrame
    changed: pd.DataFrame
    totals: pd.DataFrame
    columns_only_left: list[str] = field(default_factory=list)
    columns_only_right: list[str] = field(default_factory=list)
    duplicate_keys_left: int = 0
    duplicate_keys_right: int = 0

    @property
    def changed_rows(self) -> int:
        if self.changed.empty:
            return 0
        return len(self.changed[self.key].drop_duplicates())

    @property
    def is_identical(self) -> bool:
        return (
            self.added.empty
            and self.removed.empty
            and self.changed.empty
            and not self.columns_only_left
            and not self.columns_only_right
        )

    def summary(self) -> dict[str, object]:
        return {
            "Left file": self.left_name,
            "Right file": self.right_name,
            "Key column(s)": ", ".join(self.key),
            "Rows in left": self.left_rows,
            "Rows in right": self.right_rows,
            "Rows added": len(self.added),
            "Rows removed": len(self.removed),
            "Rows changed": self.changed_rows,
            "Cells changed": len(self.changed),
            "Columns only in left": ", ".join(self.columns_only_left) or "-",
            "Columns only in right": ", ".join(self.columns_only_right) or "-",
            "Duplicate keys (left)": self.duplicate_keys_left,
            "Duplicate keys (right)": self.duplicate_keys_right,
            "Result": "IDENTICAL" if self.is_identical else "DIFFERENCES FOUND",
        }


# --------------------------------------------------------------------------- #
# Loading
# --------------------------------------------------------------------------- #
def load_table(path: str | Path, sheet: str | int | None = None) -> pd.DataFrame:
    """Load a CSV / TSV / Excel file into a DataFrame with clean column names."""
    path = Path(path)
    if not path.exists():
        raise SheetDiffError(f"File not found: {path}")
    ext = path.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise SheetDiffError(
            f"Unsupported file type '{ext}'. Use one of: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )

    if ext in {".xlsx", ".xlsm", ".xls"}:
        try:
            df = pd.read_excel(path, sheet_name=sheet if sheet is not None else 0)
        except ValueError as exc:  # bad sheet name
            raise SheetDiffError(f"{path.name}: {exc}") from exc
    else:
        sep = "\t" if ext == ".tsv" else None
        df = pd.read_csv(path, sep=sep, engine="python")

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")  # ignore fully blank rows
    return df


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _norm_key_value(value: object) -> str:
    """Make keys comparable across CSV/Excel (e.g. 1001, 1001.0, ' 1001 ')."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _prepare(df: pd.DataFrame, key: Sequence[str]) -> tuple[pd.DataFrame, int]:
    df = df.copy()
    for k in key:
        df[k] = df[k].map(_norm_key_value)
    dupes = int(df.duplicated(subset=list(key)).sum())
    df = df.drop_duplicates(subset=list(key), keep="first")
    return df.set_index(list(key)), dupes


def _is_numeric_column(series: pd.Series) -> bool:
    non_null = series.dropna()
    if non_null.empty:
        return False
    return bool(pd.to_numeric(non_null, errors="coerce").notna().all())


def _as_text(series: pd.Series, ignore_case: bool) -> pd.Series:
    text = series.map(lambda v: "" if pd.isna(v) else str(v).strip())
    return text.str.casefold() if ignore_case else text


# --------------------------------------------------------------------------- #
# Compare
# --------------------------------------------------------------------------- #
def compare(
    left: pd.DataFrame,
    right: pd.DataFrame,
    key: str | Sequence[str],
    *,
    tolerance: float = 0.0,
    ignore_case: bool = False,
    ignore_columns: Iterable[str] = (),
    left_name: str = "left",
    right_name: str = "right",
) -> DiffResult:
    """Compare two tables row-by-row using one or more key columns.

    Args:
        left / right: the two tables (left = old/expected, right = new/actual).
        key: column(s) that uniquely identify a row.
        tolerance: numeric differences <= tolerance are treated as equal.
        ignore_case: compare text values case-insensitively.
        ignore_columns: columns to skip entirely.
    """
    key = [key] if isinstance(key, str) else list(key)
    if not key:
        raise SheetDiffError("At least one key column is required.")
    for side, df in (("left", left), ("right", right)):
        missing = [k for k in key if k not in df.columns]
        if missing:
            raise SheetDiffError(
                f"Key column(s) {missing} not found in {side} file. "
                f"Available columns: {list(df.columns)}"
            )

    ignore = set(ignore_columns)
    left_cols = [c for c in left.columns if c not in key and c not in ignore]
    right_cols = [c for c in right.columns if c not in key and c not in ignore]
    common_cols = [c for c in left_cols if c in right_cols]

    L, dup_left = _prepare(left, key)
    R, dup_right = _prepare(right, key)

    added = R.loc[~R.index.isin(L.index)].reset_index()
    removed = L.loc[~L.index.isin(R.index)].reset_index()
    common = L.index[L.index.isin(R.index)]

    # ---- cell-level changes -------------------------------------------------
    records: list[dict] = []
    for col in common_cols:
        lv = L.loc[common, col]
        rv = R.loc[common, col]
        ln = pd.to_numeric(lv, errors="coerce")
        rn = pd.to_numeric(rv, errors="coerce")
        both_num = ln.notna() & rn.notna()

        num_diff = both_num & ((rn - ln).abs() > tolerance)
        text_diff = ~both_num & (_as_text(lv, ignore_case) != _as_text(rv, ignore_case))
        mask = num_diff | text_diff

        for idx in common[mask.to_numpy()]:
            key_vals = idx if isinstance(idx, tuple) else (idx,)
            rec = dict(zip(key, key_vals))
            l_val, r_val = L.at[idx, col], R.at[idx, col]
            rec.update(
                {
                    "Column": col,
                    left_name: l_val,
                    right_name: r_val,
                    "Difference": (rn[idx] - ln[idx]) if both_num[idx] else None,
                }
            )
            records.append(rec)

    change_cols = key + ["Column", left_name, right_name, "Difference"]
    changed = pd.DataFrame(records, columns=change_cols)
    if not changed.empty:
        order = {c: i for i, c in enumerate(common_cols)}
        changed["_o"] = changed["Column"].map(order)
        changed = changed.sort_values(key + ["_o"], kind="stable").drop(columns="_o")
        changed = changed.reset_index(drop=True)

    # ---- column totals (numeric columns present in both) --------------------
    total_rows = []
    for col in common_cols:
        if _is_numeric_column(left[col]) and _is_numeric_column(right[col]):
            lt = float(pd.to_numeric(left[col], errors="coerce").sum())
            rt = float(pd.to_numeric(right[col], errors="coerce").sum())
            total_rows.append(
                {"Column": col, left_name: lt, right_name: rt, f"{right_name} − {left_name}": rt - lt}
            )
    totals = pd.DataFrame(
        total_rows, columns=["Column", left_name, right_name, f"{right_name} − {left_name}"]
    )

    return DiffResult(
        left_name=left_name,
        right_name=right_name,
        key=key,
        left_rows=len(left),
        right_rows=len(right),
        added=added,
        removed=removed,
        changed=changed,
        totals=totals,
        columns_only_left=[c for c in left_cols if c not in right_cols],
        columns_only_right=[c for c in right_cols if c not in left_cols],
        duplicate_keys_left=dup_left,
        duplicate_keys_right=dup_right,
    )


def compare_files(
    left_path: str | Path,
    right_path: str | Path,
    key: str | Sequence[str],
    *,
    left_sheet: str | int | None = None,
    right_sheet: str | int | None = None,
    **kwargs,
) -> DiffResult:
    """Load two files and compare them. Extra kwargs go to :func:`compare`."""
    left_path, right_path = Path(left_path), Path(right_path)
    left_name, right_name = left_path.stem, right_path.stem
    if left_name == right_name:
        left_name, right_name = f"{left_name} (left)", f"{right_name} (right)"
    return compare(
        load_table(left_path, left_sheet),
        load_table(right_path, right_sheet),
        key,
        left_name=left_name,
        right_name=right_name,
        **kwargs,
    )
