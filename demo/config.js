window.DEMO = {
  repo: "spreadsheet-diff",
  title: "spreadsheet-diff",
  tagline: "Drop in two versions of a CSV or Excel file. See which rows were added, removed and changed, cell by cell, keyed on the column you pick.",
  pkg: "spreadsheet_diff",
  packages: ["openpyxl"],
  inputs: [
    { id: "old", label: "Old version (CSV or XLSX)", accept: ".csv,.xlsx,.xlsm,.txt,.tsv", sample: "samples/old.xlsx" },
    { id: "new", label: "New version (CSV or XLSX)", accept: ".csv,.xlsx,.xlsm,.txt,.tsv", sample: "samples/new.csv" },
  ],
  options: [
    { id: "key", label: "Key column", type: "text", value: "id", placeholder: "id, or region,id" },
    { id: "ignore_case", label: "Ignore case", type: "checkbox", value: true },
    { id: "ignore_whitespace", label: "Ignore extra spaces", type: "checkbox", value: true },
    { id: "tolerance", label: "Number tolerance", type: "text", value: "0.001" },
  ],
  sampleNote: 'The sample is 40 synthetic customer rows: <a href="samples/old.xlsx">old.xlsx</a> and a newer <a href="samples/new.csv">new.csv</a> with one row added, one removed, a few edits, a new column and a dropped one.',
  python: `
from pathlib import Path
from spreadsheet_diff.cli import main as cli_main

def run(paths, opts):
    out = "/work/diff.xlsx"
    argv = [paths["old"][0], paths["new"][0], "--key", opts["key"].strip() or "id", "-o", out]
    if opts.get("ignore_case"):
        argv.append("--ignore-case")
    if opts.get("ignore_whitespace"):
        argv.append("--ignore-whitespace")
    if opts.get("tolerance", "").strip():
        argv += ["--tolerance", opts["tolerance"].strip()]
    code, log = run_cli(cli_main, argv)  # exit 1 means differences were found, like diff
    if code not in (0, 1) or not Path(out).exists():
        return {"ok": False, "log": log, "sheets": [], "download": None}
    return {"ok": True, "log": log, "sheets": sheets_from_xlsx(out), "download": out}
`,
};
