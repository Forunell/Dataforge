import csv
import io
import json
import re
from decimal import Decimal, InvalidOperation

import pytest
from openpyxl import Workbook, load_workbook

from dataforge.app import SESSIONS
from dataforge.core import DataError, csv_bytes, export_files, number, read_source, safe_text, suggest
from dataforge.editing import apply_edits, preview_edits
from test_app import demo_preview, token

from test_core import run, source
from test_core import DATA


@pytest.mark.parametrize("value,expected", [
    ("1 234,56", "1234.56"), ("1\u00a0234.56", "1234.56"), ("1\u202f234,56", "1234.56"),
    ("1,234.56", "1234.56"), ("1.234,56", "1234.56"), ("1,234,567", "1234567"),
    ("1.234.567", "1234567"), ("-1 234,56", "-1234.56"), ("999999999999999999", "999999999999999999"),
    ("123456789012.123456", "123456789012.123456"), ("+12.50", "12.50"),
])
def test_exact_money(value, expected):
    assert number(value) == Decimal(expected)


@pytest.mark.parametrize("value", ["1,234", "12.345", "0.001"])
def test_money_ambiguity(value):
    with pytest.raises(DataError):
        number(value)


@pytest.mark.parametrize("value", ["1,23,456", "1 23,45", "1.23.456,78", "1e20", "NaN", "Infinity",
                                      "9999999999999999999", "0.0000001", "١٢", "1,234.567,89"])
def test_money_boundaries(value):
    with pytest.raises((InvalidOperation, DataError)):
        number(value)


def test_optional_columns_and_currency_blocks_price():
    a = source([["001", "Чай", "12,50"]], headers=["Код товару", "Назва", "Ціна"])
    assert suggest(a) == dict(sku=0, name=1, price=2, currency=-1, stock=-1, category=-1, variant=-1)
    r = run(a, a)
    assert r["errors"] == 0
    assert r["comparison_counts"] == {"blocked": 1}
    assert not any(c["kind"] == "price" for c in r["changes"])
    assert "normalized.csv" not in export_files(r)
    assert r["row_counts"] == {"ok": 0, "warning": 2, "error": 0}


def test_composite_keys_and_category_name_changes():
    headers = ["sku", "name", "price", "currency", "stock", "category", "variant"]
    a = source([["001", "Чай", "10", "UAH", "1", "Напої", "S"],
                ["001", "Чай", "10", "UAH", "1", "Напої", "L"]], headers=headers)
    b = source([["001", "ЧАЙ ", "10.00", "UAH", "1", "Напої", "S"],
                ["001", "Чай", "12", "UAH", "2", "Новинки", "L"]], headers=headers)
    r = run(a, b, {"composite": True, "name_trim": True, "name_casefold": True})
    assert r["comparison_counts"] == {"unchanged": 1, "changed": 1}
    assert r["counts"].get("duplicate", 0) == 0
    assert {c["field"] for c in r["changes"]} == {"price", "stock", "category"}
    assert r["warnings"] == 0  # Exactly 20% is not above the threshold.
    assert r["changes"][0]["rows_a"] == "3"
    assert r["changes"][0]["column_b"] == "price"
    assert run(a, b)["counts"]["duplicate"] == 4


def test_zero_base_threshold_and_rule_switches():
    a = source([["001", "", "0", "UAH", "1.5"]])
    b = source([["001", "", "-2", "UAH", "1.5"]])
    r = run(a, b, {"require_name": False, "nonnegative_price": False, "integer_stock": False})
    assert r["errors"] == 0
    assert r["counts"]["threshold"] == 1
    files = export_files(r)
    parsed = list(csv.DictReader(io.StringIO(files["normalized.csv"].decode("utf-8-sig")), delimiter=";"))
    assert parsed[1]["price"] == "-2"


def test_issue_provenance_and_quality_partition():
    a = source([["", "", "-3", "UAH", "1.5"], ["D", "Товар", "3", "UAH", "1"],
                ["D", "Товар", "3", "UAH", "1"]])
    r = run(a)
    assert len(r["products"]) == 4
    assert sum(r["row_counts"].values()) == 4
    assert r["row_counts"] == {"ok": 1, "warning": 0, "error": 3}
    assert r["quality"] == 25
    issue = next(i for i in r["issues"] if i["code"] == "negative")
    assert (issue["source"], issue["file"], issue["row"], issue["column"], issue["original"]) == ("A", "test.csv", 2, "price", "-3")
    assert r["comparison_counts"] == {"blocked": 2, "new": 1}


def test_reversible_bulk_edit_preserves_original_and_ambiguous():
    a = source([[" 001 ", " чай ", "1 234,56", "UAH", "2"], ["002", "кава", "1,234", "UAH", "1"]])
    original = [r[:] for r in a.rows]
    patches = preview_edits([a, a], 0, 2, "decimal")
    assert len(patches) == 1
    edited = apply_edits([a, a], patches)
    assert a.rows == original
    assert edited[0].rows[0][2] == "1234.56"
    assert edited[0].rows[1][2] == "1,234"
    assert apply_edits(edited, patches, reverse=True)[0].rows == original
    assert preview_edits([a, a], 0, 1, "upper")[0]["after"] == " ЧАЙ "
    assert preview_edits([a, a], 0, 0, "trim")[0]["after"] == "001"
    with pytest.raises(DataError):
        apply_edits(edited, patches)


def test_xlsx_sheet_selection_empty_first_and_formula_text():
    book = Workbook()
    book.active.title = "Порожній"
    sheet = book.create_sheet("Прайс")
    sheet.append(["sku", "name", "price", "currency", "stock"])
    sheet.append(["0001", "=1+1", 10, "UAH", 1])
    buf = io.BytesIO()
    book.save(buf)
    data = buf.getvalue()
    s = read_source("many.xlsx", data, sheet_name="Прайс")
    assert s.sheets == ["Порожній", "Прайс"]
    assert s.sheet == "Прайс" and s.rows[0][1] == "=1+1"
    assert run(s)["counts"]["formula"] == 1
    with pytest.raises(DataError):
        read_source("many.xlsx", data, sheet_name="missing")


@pytest.mark.parametrize("text", [" \ttext", "\x00=1", "\x1btext", "\ufeff\ntext", "\x7ftext", " \r@cmd"])
def test_export_control_prefix(text):
    assert safe_text(text).startswith("'")


@pytest.mark.parametrize("delimiter", [",", ";", "\t"])
def test_csv_roundtrip_tabs_bom_and_safe_numbers(delimiter):
    rows = [dict(sku="000001", name='Кава; чай, "лапки"\nдругий рядок', price="-12.5", stock="2")]
    raw = csv_bytes(rows, list(rows[0]), delimiter)
    parsed = read_source("round.csv", raw, delimiter=delimiter)
    assert parsed.rows == [list(rows[0].values())]
    assert raw.startswith(b"\xef\xbb\xbf")


def test_clean_export_has_normalized_xlsx_and_report():
    a = source([["0001", "Чай", "10", "UAH", "1"]])
    r = run(a, a)
    files = export_files(r)
    assert "normalized.csv" in files and "review_rows.csv" not in files
    assert json.loads(files["report.json"])["status"] == "checked"
    book = load_workbook(io.BytesIO(files["summary.xlsx"]))
    assert book["normalized"]["D2"].value == "0001"
    assert book["normalized"]["D2"].data_type == "s"
    book.close()


def test_http_edit_apply_undo_restore_and_stale_revision(client):
    csrf, _, preview = demo_preview(client)
    rev = token(preview, "revision")
    proposed = client.post("/edit/preview", data=dict(csrf=csrf, revision=rev, target="0:2", operation="manual", row="8", value="50"))
    assert proposed.status_code == 200
    state = next(iter(SESSIONS.values()))
    assert state["sources"][0].rows[6][2] == "-50"
    applied = client.post("/edit/apply", data=dict(csrf=csrf, revision=rev, preview_token=token(proposed, "preview_token")))
    assert applied.status_code == 200
    assert state["sources"][0].rows[6][2] == "50"
    assert state["originals"][0].rows[6][2] == "-50"
    assert state["result"]["resolved_errors"] == 1
    assert state["result"]["edits"][0]["before"] == "-50"
    assert client.post("/edit/undo", data=dict(csrf=csrf, revision=rev)).status_code == 400
    undone = client.post("/edit/undo", data=dict(csrf=csrf, revision=token(applied, "revision")))
    assert undone.status_code == 200 and state["sources"][0].rows[6][2] == "-50"
    assert state["result"]["resolved_errors"] == 0
    restored = client.post("/edit/restore", data=dict(csrf=csrf, revision=token(undone, "revision")))
    assert restored.status_code == 200
    assert not state["history"] and not state["result"]["edits"]


def test_http_independent_replacement_and_invalidation(client):
    csrf, _, preview = demo_preview(client)
    client.post("/apply", data=dict(csrf=csrf, revision=token(preview, "revision"), acknowledge="on"))
    raw = b"sku;name;price;currency;stock\n0001;Tea;10;UAH;2"
    replaced = client.post("/upload", data=dict(csrf=csrf), files={"file_a": ("new.csv", raw)})
    assert replaced.status_code == 200
    state = next(iter(SESSIONS.values()))
    assert state["sources"][0].name == "new.csv"
    assert state["sources"][1].name == "supplier_b.xlsx"
    assert state["demo_sources"] == [False, True] and state["result"] is None
    assert client.get("/download/dataforge-results.zip").status_code == 404


def test_http_sheet_selection(client):
    book = Workbook()
    book.active.title = "Empty"
    sheet = book.create_sheet("Data")
    sheet.append(["sku", "name", "price"])
    sheet.append(["001", "Tea", "10"])
    buf = io.BytesIO()
    book.save(buf)
    csrf = token(client.get("/"))
    response = client.post("/upload", data=dict(csrf=csrf), files={"file_a": ("a.xlsx", buf.getvalue()), "file_b": ("b.xlsx", buf.getvalue())})
    assert response.status_code == 200 and 'name="sheet_0"' in response.text
    assert client.post("/sheets", data=dict(csrf=csrf)).status_code == 400
    assert client.post("/sheets", data=dict(csrf=csrf, sheet_0="Data", sheet_1="Data")).status_code == 200
    assert next(iter(SESSIONS.values()))["sources"][0].sheet == "Data"


def test_http_pagination_search_filter_sort_and_guards(client):
    csrf = token(client.get("/"))
    raw = "sku;name;price;currency;stock\n" + "\n".join(f"{i:04};Tea {i};10;UAH;1" for i in range(101))
    client.post("/upload", data=dict(csrf=csrf), files={"file_a": ("a.csv", raw.encode()), "file_b": ("b.csv", raw.encode())})
    mapping = {f"{s}_{k}": str(i) for s in "ab" for i, k in enumerate(["sku", "name", "price", "currency", "stock"])}
    client.post("/analyze", data=dict(csrf=csrf, compare="on", **mapping))
    page = client.get("/step/3?kind=unchanged&direction=desc&page=2")
    assert page.status_code == 200 and page.text.count("data-record") == 50
    last = client.get("/step/3?page=3")
    assert last.text.count("data-record") == 1
    assert client.get("/step/3?q=0099").text.count("data-record") == 1
    assert client.get("/step/3?kind=new").text.count("data-record") == 0
    assert client.get("/step/3?page=bad").status_code == 200
    rev = token(client.get("/step/3"), "revision")
    assert client.post("/edit/preview", data=dict(csrf=csrf, revision=rev, target="0:0", operation="manual", row="9999")).status_code == 400
    assert client.post("/analyze", data=dict(csrf=csrf, threshold="NaN", **mapping)).status_code == 400
    assert re.search(r'name="revision" value="[^"]+"', page.text)


def test_golden_fixture_results():
    golden = json.loads((DATA / "golden_results.json").read_text(encoding="utf-8"))
    for prefix, expected in golden.items():
        names = ["supplier_a.csv", "supplier_b.xlsx"] if prefix == "demo" else ["clean_a.csv", "clean_b.csv"]
        a, b = [read_source(name, (DATA / name).read_bytes()) for name in names]
        result = run(a, b)
        assert len(result["products"]) == expected["rows"]
        for key in ["errors", "warnings", "quality", "comparison_counts"]:
            assert result[key] == expected[key]
        if prefix == "demo":
            assert result["row_counts"] == expected["row_counts"]
            assert len(result["changes"]) == expected["change_events"]


def test_large_upload_remains_in_memory(client, monkeypatch):
    import tempfile

    def reject_disk(*args, **kwargs):
        raise AssertionError("Upload must not spill onto disk")

    monkeypatch.setattr(tempfile, "TemporaryFile", reject_disk)
    csrf = token(client.get("/"))
    raw = ("sku;name;price;currency;stock\n" + "\n".join(f"{i};{'x'*900};10;UAH;1" for i in range(1500))).encode()
    assert len(raw) > 1024 * 1024
    response = client.post("/upload", data=dict(csrf=csrf), files={"file_a": ("a.csv", raw), "file_b": ("b.csv", raw)})
    assert response.status_code == 200
    assert len(next(iter(SESSIONS.values()))["sources"][0].rows) == 1500
