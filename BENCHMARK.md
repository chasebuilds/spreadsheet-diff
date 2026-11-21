# Benchmark

Written by `python -m benchmark.run_benchmark --seed 11 --rows 20000` on
2026-11-21 (3.11.15, Windows). Every number
below comes from that run. Do not edit by hand.

## Method

`benchmark/generate.py` makes 20,000 fake customer rows (id, name, email, city, amount, qty,
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

| Measure | Result |
|---|---|
| Rows (old file .xlsx, new file .csv) | 20,000 old, 20,030 new |
| Added rows | recall **100.0%** (600 of 600), precision **100.0%** (600 of 600) |
| Removed rows | recall **100.0%** (570 of 570), precision **100.0%** (570 of 570) |
| Changed rows | recall **100.0%** (1,377 of 1,377), precision **100.0%** (1,377 of 1,377) |
| Changed cells | recall **100.0%** (2,803 of 2,803), precision **100.0%** (2,803 of 2,803) |
| Noise rows flagged, with ignore options | 0 of 1,172 |
| Noise rows flagged, strict (no options) | 1,172 of 1,172 |
| Column added, removed and type change found | all 3 |
| Time to load both files and diff | 1.8 s |

Strict mode is the same diff with no ignore options. It shows what the options remove: every
noise row is a change when case, spaces and tiny number gaps count.

## What this does not show

The planted edits are clean and the key column is never edited, so a real file with renamed
keys, merged cells or several header rows will do worse. See [LIMITATIONS.md](LIMITATIONS.md).
