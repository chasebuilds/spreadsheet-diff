"""Make an old and a new table with known planted changes.

Planted: removed rows, added rows, rows with 1 to 3 real cell edits, a dropped column, a new
column, and a column whose type changes (whole numbers written as decimals). Noise rows get
edits that should not count when the ignore options are on: case only, extra spaces only, or a
number off by less than 0.001.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import random
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from spreadsheet_diff.diff import Table

FIRST = ["Ava", "Ben", "Cara", "Dev", "Ezra", "Faye", "Gus", "Hana", "Ivo", "June", "Kai", "Lena",
         "Milo", "Nia", "Otto", "Pia", "Quin", "Rosa", "Seth", "Tess", "Uma", "Vic", "Wren", "Yara"]
LAST = ["Abbot", "Brook", "Crane", "Dale", "Ellis", "Frost", "Grove", "Hale", "Irwin", "Jolly",
        "Knox", "Lowe", "Marsh", "North", "Oakes", "Pike", "Reed", "Stone", "Tate", "Vale"]
CITIES = ["Ashby", "Brookvale", "Cedar Falls", "Dunmore", "Eastport", "Fairhaven", "Glenwood",
          "Harlow", "Ironton", "Juniper"]
STATUS = ["active", "paused", "closed", "trial"]
EDITABLE = ["name", "email", "city", "amount", "qty", "joined", "status"]
OLD_HEADERS = ["id", "name", "email", "city", "amount", "qty", "joined", "status", "region"]
NEW_HEADERS = ["id", "name", "email", "city", "amount", "qty", "joined", "status", "tier"]


@dataclass
class Case:
    old: Table
    new: Table
    added: set[str] = field(default_factory=set)
    removed: set[str] = field(default_factory=set)
    cells: set[tuple[str, str]] = field(default_factory=set)
    noise: set[str] = field(default_factory=set)


def _row(rng: random.Random, i: int) -> dict[str, Any]:
    first, last = rng.choice(FIRST), rng.choice(LAST)
    return {
        "id": i,
        "name": f"{first} {last}",
        "email": f"{first.lower()}.{last.lower()}{rng.randint(1, 99)}@example.com",
        "city": rng.choice(CITIES),
        "amount": round(rng.uniform(5, 5000), 2),
        "qty": rng.randint(1, 50),
        "joined": dt.date(2020, 1, 1) + dt.timedelta(days=rng.randint(0, 2000)),
        "status": rng.choice(STATUS),
        "region": rng.choice(["east", "west"]),
    }


def _edit(rng: random.Random, col: str, v: Any) -> Any:
    if col == "name":
        while (n := f"{rng.choice(FIRST)} {rng.choice(LAST)}").casefold() == v.casefold():
            pass
        return n
    if col == "email":
        return "x" + v
    if col == "city":
        return rng.choice([c for c in CITIES if c != v])
    if col == "amount":
        return round(v + rng.uniform(1, 500), 2)
    if col == "qty":
        return v + rng.randint(1, 5)
    if col == "joined":
        return v + dt.timedelta(days=rng.randint(1, 90))
    return rng.choice([s for s in STATUS if s != v])


def _noise(rng: random.Random, row: dict[str, Any]) -> None:
    pick = rng.randrange(3)
    if pick == 0:
        row["name"] = row["name"].upper()
    elif pick == 1:
        row["city"] = f"  {row['city']} "
    else:
        row["amount"] = round(row["amount"] + rng.choice([-1, 1]) * 0.0004, 4)


def generate(seed: int, rows: int) -> Case:
    rng = random.Random(seed)
    old = [_row(rng, i) for i in range(1, rows + 1)]
    case = Case(Table(OLD_HEADERS, old), Table(NEW_HEADERS, []))
    new: list[dict[str, Any]] = []
    for r in old:
        roll = rng.random()
        if roll < 0.03:
            case.removed.add(str(r["id"]))
            continue
        n = {h: r.get(h) for h in NEW_HEADERS}
        n["tier"] = rng.choice(["A", "B", "C"])
        n["qty"] = float(n["qty"])
        if roll < 0.10:
            for col in rng.sample(EDITABLE, rng.randint(1, 3)):
                n[col] = _edit(rng, col, r[col])
                case.cells.add((str(r["id"]), col))
        elif roll < 0.16:
            _noise(rng, n)
            case.noise.add(str(r["id"]))
        new.append(n)
    for i in range(rows + 1, rows + 1 + max(1, rows * 3 // 100)):
        n = {h: v for h, v in _row(rng, i).items() if h in NEW_HEADERS}
        n["tier"] = "A"
        new.append(n)
        case.added.add(str(i))
    rng.shuffle(new)
    case.new.rows = new
    return case


def write(case: Case, out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.append(case.old.headers)
    for r in case.old.rows:
        ws.append([r[h] for h in case.old.headers])
    wb.save(out / "old.xlsx")
    with (out / "new.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(case.new.headers)
        for r in case.new.rows:
            w.writerow([r[h] for h in case.new.headers])


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--seed", type=int, default=7)
    p.add_argument("--rows", type=int, default=40)
    p.add_argument("--out", default="samples")
    a = p.parse_args()
    write(generate(a.seed, a.rows), Path(a.out))
