# Limitations

What spreadsheet-diff does not do, or does poorly. Read this before trusting a report.

## Matching rows

- **Rows are matched by key only.** If the key itself was edited, the row shows up once as removed
  and once as added. There is no fuzzy or similarity matching.
- **Duplicate keys keep the first row.** Later rows with the same key are not compared. They are
  listed under duplicate keys in the Summary sheet and the JSON.
- **Row order is ignored.** A row that only moved is not reported.
- **Keys are compared as text after trimming.** `1`, `1.0` (from Excel) and ` 1 ` match, but
  `001` and `1` do not.

## Comparing values

- **Anything that reads as a number is compared as a number.** `007` and `7` are equal. If a code
  column like a ZIP or a part number must keep its leading zeros, the change will not show as a
  cell change, only as a column type change when the whole column changes.
- **Numbers with commas or currency signs are text.** `1,200` and `1200` are different.
- **Only ISO dates are read from text.** `2026-01-02` matches an Excel date; `1/2/2026` does not.
- **`--ignore-whitespace` collapses all spaces**, including ones inside the text that might matter.
- **Formulas are read as their last saved value.** A workbook saved by a tool that does not store
  values (some scripts) reads as empty cells.

## Files

- **One sheet per file**, the first or the one named with `--sheet`. The same sheet name is used for
  both files.
- **The first row must be the header.** Merged cells, several header rows or a title row above the
  table give wrong column names.
- **CSV files must be UTF-8.** Other encodings fail to load.
- **Whole files are loaded into memory.** A few hundred thousand rows is fine on a normal laptop;
  millions of rows may not be.
- **Columns are matched by exact header text.** A renamed column shows as one removed and one added.
