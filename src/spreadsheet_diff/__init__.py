"""Compare two versions of a CSV or Excel file by key."""

from spreadsheet_diff.diff import CellChange, Diff, Options, Table, compare, load

__all__ = ["CellChange", "Diff", "Options", "Table", "compare", "load"]
__version__ = "0.1.0"
