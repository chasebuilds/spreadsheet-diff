"""Write the Excel report and the JSON summary."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill

from spreadsheet_diff.diff import Diff, Table, jsonable, text

FILL = PatternFill("solid", start_color="FFF2A8", end_color="FFF2A8")
BOLD = Font(bold=True)


def _sheet(wb: Any, title: str, headers: list[str]) -> Any:
    ws = wb.create_sheet(title)
    ws.append(headers)
    for c in ws[1]:
        c.font = BOLD
    ws.freeze_panes = "A2"
    return ws


def _cell(v: Any) -> Any:
    return v if v is None or isinstance(v, int | float | str) else text(v)


def write_excel(d: Diff, old: Table, new: Table, path: str | Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    s = d.summary()
    ws.append(["Measure", "Value"])
    for c in ws[1]:
        c.font = BOLD
    for label, value in [
        ("Key", ", ".join(d.key)),
        ("Rows in old file", d.rows_old),
        ("Rows in new file", d.rows_new),
        ("Rows added", s["rows_added"]),
        ("Rows removed", s["rows_removed"]),
        ("Rows changed", s["rows_changed"]),
        ("Cells changed", s["cells_changed"]),
        ("Columns added", ", ".join(d.columns_added)),
        ("Columns removed", ", ".join(d.columns_removed)),
        ("Type changes", ", ".join(f"{c}: {o} to {n}" for c, (o, n) in d.type_changes.items())),
        ("Duplicate keys in old file", ", ".join(d.duplicate_keys_old)),
        ("Duplicate keys in new file", ", ".join(d.duplicate_keys_new)),
    ]:
        ws.append([label, value])
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 40

    added = _sheet(wb, "Added", new.headers)
    for r in d.added:
        added.append([_cell(r.get(h)) for h in new.headers])
    removed = _sheet(wb, "Removed", old.headers)
    for r in d.removed:
        removed.append([_cell(r.get(h)) for h in old.headers])

    changed = _sheet(wb, "Changed", new.headers)
    for k, cells in d.changed.items():
        row = d.new_rows[k]
        changed.append([_cell(row.get(h)) for h in new.headers])
        olds = {c.column: c.old for c in cells}
        for i, h in enumerate(new.headers, start=1):
            if h in olds:
                cell = changed.cell(row=changed.max_row, column=i)
                cell.fill = FILL
                cell.comment = Comment(f"old: {text(olds[h])}", "sdiff")

    cells_ws = _sheet(wb, "Cells", ["key", "column", "old", "new"])
    for cells in d.changed.values():
        for c in cells:
            cells_ws.append([c.key, c.column, _cell(c.old), _cell(c.new)])
    wb.save(path)


def write_json(d: Diff, path: str | Path) -> None:
    Path(path).write_text(json.dumps(d.summary(), indent=2, default=jsonable), encoding="utf-8")
