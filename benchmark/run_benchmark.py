"""Diff generated files with planted changes, score them, write BENCHMARK.md and README table."""

from __future__ import annotations

import argparse
import datetime as dt
import platform
import re
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path

from benchmark.generate import generate, write
from spreadsheet_diff.diff import Options, compare, key_of, load

ROOT = Path(__file__).resolve().parent.parent
OPTS = Options(ignore_case=True, ignore_whitespace=True, tolerance=0.001)


@dataclass
class Score:
    rows_old: int
    rows_new: int
    added: tuple[int, int, int]  # (true, found, right)
    removed: tuple[int, int, int]
    cells: tuple[int, int, int]
    changed_rows: tuple[int, int, int]
    noise: int
    noise_flagged: int
    noise_flagged_strict: int
    columns_ok: bool
    seconds: float


def _prf(true: set[str] | set[tuple[str, str]], found: set[str] | set[tuple[str, str]]) -> tuple[
        int, int, int]:
    return len(true), len(found), len(true & found)  # type: ignore[operator]


def score(seed: int, rows: int) -> Score:
    case = generate(seed, rows)
    with tempfile.TemporaryDirectory() as tmp:
        write(case, Path(tmp))
        start = time.perf_counter()
        old, new = load(Path(tmp) / "old.xlsx"), load(Path(tmp) / "new.csv")
        d = compare(old, new, ["id"], OPTS)
        seconds = time.perf_counter() - start
        strict = compare(old, new, ["id"])
    cells = {(c.key, c.column) for cs in d.changed.values() for c in cs}
    return Score(
        rows_old=len(old.rows), rows_new=len(new.rows),
        added=_prf(case.added, {key_of(r, ["id"]) for r in d.added}),
        removed=_prf(case.removed, {key_of(r, ["id"]) for r in d.removed}),
        cells=_prf(case.cells, cells),
        changed_rows=_prf({k for k, _ in case.cells}, set(d.changed)),
        noise=len(case.noise),
        noise_flagged=len(case.noise & set(d.changed)),
        noise_flagged_strict=len(case.noise & set(strict.changed)),
        columns_ok=(d.columns_added == ["tier"] and d.columns_removed == ["region"]
                    and d.type_changes == {"qty": ("int", "float")}),
        seconds=seconds,
    )


def _pr(t: tuple[int, int, int]) -> str:
    true, found, right = t
    rec = f"{100 * right / true:.1f}%" if true else "n/a"
    prec = f"{100 * right / found:.1f}%" if found else "n/a"
    return (f"recall **{rec}** ({right:,} of {true:,}), "
            f"precision **{prec}** ({right:,} of {found:,})")


def summary_table(s: Score) -> str:
    return "\n".join([
        "| Measure | Result |",
        "|---|---|",
        f"| Rows (old file .xlsx, new file .csv) | {s.rows_old:,} old, {s.rows_new:,} new |",
        f"| Added rows | {_pr(s.added)} |",
        f"| Removed rows | {_pr(s.removed)} |",
        f"| Changed rows | {_pr(s.changed_rows)} |",
        f"| Changed cells | {_pr(s.cells)} |",
        f"| Noise rows flagged, with ignore options | {s.noise_flagged} of {s.noise:,} |",
        f"| Noise rows flagged, strict (no options) | {s.noise_flagged_strict:,} of {s.noise:,} |",
        f"| Column added, removed and type change found | {'all 3' if s.columns_ok else 'no'} |",
        f"| Time to load both files and diff | {s.seconds:.1f} s |",
    ])


def report(s: Score, seed: int, rows: int) -> str:
    return f"""# Benchmark

Written by `python -m benchmark.run_benchmark --seed {seed} --rows {rows}` on
{dt.date.today().isoformat()} ({platform.python_version()}, {platform.system()}). Every number
below comes from that run. Do not edit by hand.

## Method

`benchmark/generate.py` makes {rows:,} fake customer rows (id, name, email, city, amount, qty,
joined date, status, region) with a fixed seed, saves them as `old.xlsx`, then makes a new
version saved as `new.csv` with known planted changes:

- about 3% of rows removed and 3% new rows added,
- about 7% of rows with 1 to 3 real edits (new name, email, city, date or status, an amount up by
  1 to 500, a qty up by 1 to 5),
- about 6% of rows with noise only: name in upper case, extra spaces around the city, or an
  amount off by 0.0004,
- column `region` dropped, column `tier` added, and `qty` written as `3.0` instead of `3`.

The tool loads both files from disk and diffs them by `id` with `--ignore-case
--ignore-whitespace --tolerance 0.001`. Each result is checked against the planted list.
A row or cell counts as right only if its key (and column) matches a planted one.

## Result

{summary_table(s)}

Strict mode is the same diff with no ignore options. It shows what the options remove: every
noise row is a change when case, spaces and tiny number gaps count.

## What this does not show

The planted edits are clean and the key column is never edited, so a real file with renamed
keys, merged cells or several header rows will do worse. See [LIMITATIONS.md](LIMITATIONS.md).
"""


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--seed", type=int, default=11)
    p.add_argument("--rows", type=int, default=20000)
    a = p.parse_args()
    s = score(a.seed, a.rows)
    (ROOT / "BENCHMARK.md").write_text(report(s, a.seed, a.rows), encoding="utf-8")
    readme = ROOT / "README.md"
    text = readme.read_text(encoding="utf-8")
    text = re.sub(r"<!-- benchmark:start -->.*<!-- benchmark:end -->",
                  f"<!-- benchmark:start -->\n{summary_table(s)}\n<!-- benchmark:end -->",
                  text, flags=re.S)
    readme.write_text(text, encoding="utf-8")
    print(summary_table(s))


if __name__ == "__main__":
    main()
