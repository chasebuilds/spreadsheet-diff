"""Load tables and compare them by key."""

from __future__ import annotations

import csv
import datetime as dt
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

INT_RE = re.compile(r"^-?(0|[1-9]\d*)$")
FLOAT_RE = re.compile(r"^-?\d*\.\d+([eE][-+]?\d+)?$|^-?\d+[eE][-+]?\d+$")
PADDED_RE = re.compile(r"^-?0\d+$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$")


@dataclass
class Table:
    headers: list[str]
    rows: list[dict[str, Any]]


@dataclass
class Options:
    ignore_case: bool = False
    ignore_whitespace: bool = False
    tolerance: float = 0.0


@dataclass
class CellChange:
    key: str
    column: str
    old: Any
    new: Any


@dataclass
class Diff:
    key: list[str]
    columns_added: list[str]
    columns_removed: list[str]
    type_changes: dict[str, tuple[str, str]]
    added: list[dict[str, Any]]
    removed: list[dict[str, Any]]
    changed: dict[str, list[CellChange]]
    new_rows: dict[str, dict[str, Any]]
    duplicate_keys_old: list[str] = field(default_factory=list)
    duplicate_keys_new: list[str] = field(default_factory=list)
    rows_old: int = 0
    rows_new: int = 0

    @property
    def same(self) -> bool:
        return not (self.added or self.removed or self.changed or self.columns_added
                    or self.columns_removed or self.type_changes)

    def summary(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "rows_old": self.rows_old,
            "rows_new": self.rows_new,
            "rows_added": len(self.added),
            "rows_removed": len(self.removed),
            "rows_changed": len(self.changed),
            "cells_changed": sum(len(c) for c in self.changed.values()),
            "columns_added": self.columns_added,
            "columns_removed": self.columns_removed,
            "type_changes": {c: {"old": o, "new": n} for c, (o, n) in self.type_changes.items()},
            "duplicate_keys_old": self.duplicate_keys_old,
            "duplicate_keys_new": self.duplicate_keys_new,
            "added_keys": [key_of(r, self.key) for r in self.added],
            "removed_keys": [key_of(r, self.key) for r in self.removed],
            "changed": {k: {c.column: {"old": jsonable(c.old), "new": jsonable(c.new)} for c in v}
                        for k, v in self.changed.items()},
        }


def jsonable(v: Any) -> Any:
    return v.isoformat() if isinstance(v, dt.date | dt.time) else v


def load(path: str | Path, sheet: str | None = None) -> Table:
    p = Path(path)
    raw: list[list[Any]]
    if p.suffix.lower() in {".xlsx", ".xlsm"}:
        from openpyxl import load_workbook

        wb = load_workbook(p, read_only=True, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    else:
        with p.open(newline="", encoding="utf-8-sig") as f:
            raw = [list(r) for r in csv.reader(f)]
    if not raw:
        return Table([], [])
    headers = [str(h).strip() if h is not None else f"column_{i + 1}" for i, h in enumerate(raw[0])]
    rows = []
    for r in raw[1:]:
        if all(v is None or v == "" for v in r):
            continue
        rows.append({h: (r[i] if i < len(r) else None) for i, h in enumerate(headers)})
    return Table(headers, rows)


def kind(v: Any) -> str | None:
    """Type of one cell: int, float, bool, date, text, or None when empty."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, int):
        return "int"
    if isinstance(v, float):
        return "float"
    if isinstance(v, dt.date | dt.time):
        return "date"
    s = str(v).strip()
    if INT_RE.match(s):
        return "int"
    if FLOAT_RE.match(s):
        return "float"
    if DATE_RE.match(s):
        return "date"
    if s.lower() in {"true", "false"}:
        return "bool"
    return "text"


def column_type(values: list[Any]) -> str | None:
    kinds = Counter(k for v in values if (k := kind(v)))
    if not kinds:
        return None
    if set(kinds) <= {"int", "float"}:
        return "float" if "float" in kinds else "int"
    return next(iter(kinds)) if len(kinds) == 1 else "mixed"


def text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, dt.datetime) and v.time() == dt.time():
        return v.date().isoformat()
    if isinstance(v, dt.date | dt.time):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def number(v: Any) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, int | float):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if INT_RE.match(s) or FLOAT_RE.match(s) or PADDED_RE.match(s):
            return float(s)
    return None


def norm(v: Any, opt: Options) -> str:
    s = text(v)
    if opt.ignore_whitespace:
        s = " ".join(s.split())
    if opt.ignore_case:
        s = s.casefold()
    return s


def equal(a: Any, b: Any, opt: Options) -> bool:
    na, nb = number(a), number(b)
    if na is not None and nb is not None:
        return abs(na - nb) <= opt.tolerance
    return norm(a, opt) == norm(b, opt)


def key_of(row: dict[str, Any], key: list[str]) -> str:
    return "|".join(text(row.get(k)).strip() for k in key)


Row = dict[str, Any]


def index(rows: list[Row], key: list[str]) -> tuple[dict[str, Row], list[str]]:
    out: dict[str, dict[str, Any]] = {}
    dups: set[str] = set()
    for r in rows:
        k = key_of(r, key)
        if k in out:
            dups.add(k)
        else:
            out[k] = r
    return out, sorted(dups)


def compare(old: Table, new: Table, key: list[str], opt: Options | None = None) -> Diff:
    opt = opt or Options()
    for name, t in (("old", old), ("new", new)):
        missing = [k for k in key if k not in t.headers]
        if missing:
            raise KeyError(f"key column {', '.join(missing)} not in {name} file")
    a, dup_a = index(old.rows, key)
    b, dup_b = index(new.rows, key)
    common = [h for h in new.headers if h in old.headers and h not in key]
    type_changes = {}
    for h in common:
        ta, tb = column_type([r[h] for r in old.rows]), column_type([r[h] for r in new.rows])
        if ta and tb and ta != tb:
            type_changes[h] = (ta, tb)
    changed: dict[str, list[CellChange]] = {}
    for k, rb in b.items():
        ra = a.get(k)
        if ra is None:
            continue
        cells = [CellChange(k, h, ra[h], rb[h]) for h in common if not equal(ra[h], rb[h], opt)]
        if cells:
            changed[k] = cells
    return Diff(
        key=key,
        columns_added=[h for h in new.headers if h not in old.headers],
        columns_removed=[h for h in old.headers if h not in new.headers],
        type_changes=type_changes,
        added=[r for k, r in b.items() if k not in a],
        removed=[r for k, r in a.items() if k not in b],
        changed=changed,
        new_rows=b,
        duplicate_keys_old=dup_a,
        duplicate_keys_new=dup_b,
        rows_old=len(old.rows),
        rows_new=len(new.rows),
    )
