"""Compare two versions of a CSV or Excel file by a key column."""

from __future__ import annotations

import argparse
from pathlib import Path

from spreadsheet_diff.diff import Options, compare, load
from spreadsheet_diff.report import write_excel, write_json


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="sdiff", description=__doc__)
    p.add_argument("old")
    p.add_argument("new")
    p.add_argument("--key", required=True, help="key column, or several joined by commas")
    p.add_argument("-o", "--output", default="diff.xlsx", help="Excel report (default diff.xlsx)")
    p.add_argument("--json", help="also write a JSON summary to this file")
    p.add_argument("--sheet", help="sheet to read from Excel files (default: the first)")
    p.add_argument("--ignore-case", action="store_true")
    p.add_argument("--ignore-whitespace", action="store_true",
                   help="ignore leading, trailing and repeated spaces")
    p.add_argument("--tolerance", type=float, default=0.0,
                   help="numbers this close count as equal (default 0)")
    a = p.parse_args(argv)
    key = [k.strip() for k in a.key.split(",") if k.strip()]
    try:
        old, new = load(a.old, a.sheet), load(a.new, a.sheet)
        d = compare(old, new, key, Options(a.ignore_case, a.ignore_whitespace, a.tolerance))
    except (OSError, KeyError, ValueError) as e:
        print(f"error: {e}")
        return 2
    s = d.summary()
    print(f"{d.rows_old} rows old, {d.rows_new} rows new: {s['rows_added']} added, "
          f"{s['rows_removed']} removed, {s['rows_changed']} changed "
          f"({s['cells_changed']} cells)")
    if d.columns_added or d.columns_removed or d.type_changes:
        types = ", ".join(f"{c} {o} to {n}" for c, (o, n) in d.type_changes.items())
        print(f"columns added: {', '.join(d.columns_added) or '-'}; "
              f"removed: {', '.join(d.columns_removed) or '-'}; type changes: {types or '-'}")
    if d.duplicate_keys_old or d.duplicate_keys_new:
        print(f"duplicate keys (first row kept): {len(d.duplicate_keys_old)} old, "
              f"{len(d.duplicate_keys_new)} new")
    write_excel(d, old, new, a.output)
    print(f"Wrote {Path(a.output)}")
    if a.json:
        write_json(d, a.json)
        print(f"Wrote {Path(a.json)}")
    return 0 if d.same else 1
