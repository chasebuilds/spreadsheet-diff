import datetime as dt
import json

from openpyxl import Workbook, load_workbook

from spreadsheet_diff import Options, Table, compare, load
from spreadsheet_diff.cli import main
from spreadsheet_diff.diff import column_type

OLD = Table(["id", "name", "amount", "zip", "gone"], [
    {"id": 1, "name": "Ann", "amount": 10.0, "zip": 2134, "gone": "x"},
    {"id": 2, "name": "Bo", "amount": 5.5, "zip": 10001, "gone": "y"},
    {"id": 3, "name": "Cy", "amount": 1.0, "zip": 60601, "gone": "z"},
])
NEW = Table(["id", "name", "amount", "zip", "extra"], [
    {"id": "1", "name": "ann ", "amount": "10.000001", "zip": "02134", "extra": 1},
    {"id": "2", "name": "Bob", "amount": "5.5", "zip": "10001", "extra": 2},
    {"id": "4", "name": "Di", "amount": "3", "zip": "94105", "extra": 3},
])


def test_rows_and_columns() -> None:
    d = compare(OLD, NEW, ["id"])
    assert [r["id"] for r in d.added] == ["4"]
    assert [r["id"] for r in d.removed] == [3]
    assert d.columns_added == ["extra"] and d.columns_removed == ["gone"]
    assert set(d.changed) == {"1", "2"}
    assert {c.column for c in d.changed["1"]} == {"name", "amount"}
    assert d.type_changes == {"zip": ("int", "mixed")}


def test_options_hide_noise() -> None:
    d = compare(OLD, NEW, ["id"], Options(ignore_case=True, ignore_whitespace=True,
                                          tolerance=0.001))
    assert set(d.changed) == {"2"}
    assert [(c.column, c.old, c.new) for c in d.changed["2"]] == [("name", "Bo", "Bob")]


def test_duplicate_keys_and_missing_key() -> None:
    t = Table(["id", "v"], [{"id": 1, "v": "a"}, {"id": 1, "v": "b"}])
    assert compare(t, t, ["id"]).duplicate_keys_old == ["1"]
    try:
        compare(t, Table(["v"], []), ["id"])
    except KeyError as e:
        assert "new" in str(e)
    else:
        raise AssertionError("expected KeyError")


def test_column_type() -> None:
    assert column_type([1, "2", None]) == "int"
    assert column_type(["1.5", 2]) == "float"
    assert column_type(["a", 1]) == "mixed"
    assert column_type([dt.date(2026, 1, 2)]) == "date"
    assert column_type(["", None]) is None


def test_cli_end_to_end(tmp_path) -> None:  # type: ignore[no-untyped-def]
    old = tmp_path / "old.csv"
    old.write_text("id,name,qty\n1,Ann,2\n2,Bo,3\n", encoding="utf-8")
    wb = Workbook()
    ws = wb.active
    for r in [["id", "name", "qty"], [1, "Ann", 2], [2, "Bo", 4], [3, "Cy", 1]]:
        ws.append(r)
    new = tmp_path / "new.xlsx"
    wb.save(new)
    out, js = tmp_path / "d.xlsx", tmp_path / "d.json"
    assert main([str(old), str(new), "--key", "id", "-o", str(out), "--json", str(js)]) == 1
    summary = json.loads(js.read_text(encoding="utf-8"))
    assert summary["rows_added"] == 1
    assert summary["changed"] == {"2": {"qty": {"old": "3", "new": 4}}}
    rep = load_workbook(out)
    assert rep.sheetnames == ["Summary", "Added", "Removed", "Changed", "Cells"]
    cell = rep["Changed"]["C2"]
    assert cell.value == 4 and cell.fill.start_color.rgb.endswith("FFF2A8")
    assert cell.comment.text == "old: 3"
    assert main([str(old), str(old), "--key", "id", "-o", str(out)]) == 0
    assert main([str(old), str(new), "--key", "nope", "-o", str(out)]) == 2
    assert len(load(new).rows) == 3
