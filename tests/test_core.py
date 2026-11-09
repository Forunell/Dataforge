import csv
import io
import json
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from dataforge.core import (FIELDS, MAX_BYTES, DataError, analyze, export_files, read_source,
                            safe_text, suggest)

DATA = Path(__file__).resolve().parents[1] / "TEST_DATA"
HEADERS = ["sku", "name", "price", "currency", "stock"]


def source(rows, name="test.csv", delimiter=";", encoding="utf-8-sig", headers=HEADERS):
    stream = io.StringIO(newline="")
    writer = csv.writer(stream, delimiter=delimiter)
    writer.writerow(headers)
    writer.writerows(rows)
    return read_source(name, stream.getvalue().encode(encoding), encoding=encoding)


def run(a, b=None, rules=None, compare=True):
    if b is None:
        b = source([["OTHER", "Інший", "1", "UAH", "0"]])
    return analyze([a, b], [suggest(a), suggest(b)], rules or {}, compare)


@pytest.mark.parametrize("encoding", ["utf-8-sig", "utf-8", "cp1251"])
@pytest.mark.parametrize("delimiter", [",", ";"])
def test_csv_encodings_delimiters(encoding, delimiter):
    row = ["000123", 'Текст, "лапки"\nновий рядок', "12,50", "UAH", "2"]
    stream = io.StringIO(newline="")
    writer = csv.writer(stream, delimiter=delimiter)
    writer.writerow(HEADERS)
    writer.writerow(row)
    actual = read_source("a.csv", stream.getvalue().encode(encoding), "cp1251" if encoding == "cp1251" else "auto")
    assert actual.rows == [row]
    assert actual.delimiter == delimiter


def test_legacy_requires_choice():
    with pytest.raises(DataError, match="1251"):
        read_source("a.csv", (DATA / "cp1251.csv").read_bytes())


def test_manual_delimiter():
    result = read_source("a.csv", b"sku;name;price;currency;stock\n001;x;1;UAH;2", delimiter=";")
    assert result.rows[0][0] == "001"


@pytest.mark.parametrize("file", ["broken.csv", "broken.xlsx", "empty.csv", "oversize.csv", "too_many_rows.csv"])
def test_invalid_fixtures(file):
    with pytest.raises(DataError):
        read_source(file, (DATA / file).read_bytes(), delimiter=";")


@pytest.mark.parametrize("payload", [b"a;a\n1;2", b"a;\n1;2", b"a;b\n1;2;3", b"a;b\n1;\x00"])
def test_malformed_structure(payload):
    with pytest.raises(DataError):
        read_source("a.csv", payload, delimiter=";")


def test_limits():
    with pytest.raises(DataError):
        read_source("big.csv", b"x" * (MAX_BYTES + 1))
    with pytest.raises(DataError):
        source([["A", "x" * 1001, "1", "UAH", "1"]])
    with pytest.raises(DataError):
        read_source("bad.xls", b"123")


def test_blank_rows_never_disappear():
    a = read_source("blank.csv", b"sku;name;price;currency;stock\nA;x;1;UAH;2\n\n", delimiter=";")
    result = run(a)
    assert len(result["products"]) == 3
    assert sum(i["code"] == "required" for i in result["issues"]) == 5


def test_sku_changes_opt_in_and_collision():
    a = source([["  ab  01  ", "Товар", "1", "UAH", "1"], ["AB 01", "Товар", "2", "UAH", "2"]])
    untouched = run(a, rules={"trim": True})
    assert untouched["products"][0]["sku"] == "  ab  01  "
    normalized = run(a, rules={"sku_trim": True, "sku_spaces": True, "sku_upper": True})
    assert normalized["products"][0]["sku"] == "AB 01"
    assert normalized["counts"]["collision"] == 2
    assert normalized["counts"]["duplicate"] == 2
    assert next(c for c in normalized["changes"] if c["sku"] == "AB 01")["kind"] == "blocked"


@pytest.mark.parametrize("value,field,code", [("-1", "price", "negative"), ("-2", "stock", "negative"),
    ("1.2", "stock", "integer"), ("NaN", "price", "number"), ("1e3", "price", "number"),
    ("1,234", "price", "ambiguous"), ("2026-11-05", "price", "number"), ("", "price", "required")])
def test_invalid_numbers(value, field, code):
    row = ["0001", "Товар", "1", "UAH", "2"]
    row[HEADERS.index(field)] = value
    result = run(source([row]))
    assert any(i["field"] == field and i["code"] == code for i in result["issues"])


def test_currency_comparison():
    a = source([["0001", "Товар", "10", "USD", "2"]])
    b = source([["0001", "Товар", "999", "EUR", "3"]])
    result = run(a, b)
    assert {c["kind"] for c in result["changes"]} == {"currency", "stock"}
    assert "неможливе без курсу" in result["changes"][0]["message"]


def test_comparison_new_removed_equal_decimal_and_identity():
    a = source([["EQ", "Товар", "10,00", "UAH", "2"], ["OLD", "Старий", "1", "UAH", "1"], ["ID", "Один", "1", "UAH", "1"]])
    b = source([["EQ", "Товар", "10", "UAH", "2"], ["NEW", "Новий", "1", "UAH", "1"], ["ID", "Інший", "2", "UAH", "1"]])
    result = run(a, b)
    assert [(c["sku"], c["kind"]) for c in result["changes"]] == [("OLD", "removed"), ("ID", "price"), ("ID", "name"), ("NEW", "new")]
    assert result["comparison_counts"] == {"unchanged": 1, "removed": 1, "changed": 1, "new": 1}
    assert run(a, b, compare=False)["changes"] == []


def test_ambiguous_mapping():
    a = source([["1", "2", "Товар", "1", "UAH", "1"]], headers=["sku", "Артикул", "name", "price", "currency", "stock"])
    assert suggest(a)["sku"] == -1
    with pytest.raises(DataError):
        run(a)


def test_duplicate_mapping_invalid():
    a = source([["A", "Товар", "1", "UAH", "1"]])
    with pytest.raises(DataError):
        analyze([a, a], [{k: 0 for k in FIELDS}, suggest(a)], {})


@pytest.mark.parametrize("value", ["=1+1", "+cmd", "-cmd", "@SUM(A1)", "\t=2", "\r=2", "\n@x", "  =2", "\ufeff=2"])
def test_formula_sanitization(value):
    assert safe_text(value).startswith("'")
    assert safe_text("000123") == "000123"


def test_exports_roundtrip_and_no_row_loss():
    a = source([["000123", 'Назва, "лапки"\nдругий рядок', "12,50", "UAH", "2"], ["=bad", "@cmd", "-1", "UAH", "1"]])
    result = run(a)
    files = export_files(result)
    assert "normalized.csv" not in files
    normalized = list(csv.DictReader(io.StringIO(files["review_rows.csv"].decode("utf-8-sig")), delimiter=";"))
    assert len(normalized) == 3
    assert normalized[0]["sku"] == "000123"
    assert normalized[0]["name"] == a.rows[0][1]
    assert normalized[1]["sku"] == "'=bad"
    assert json.loads(normalized[1]["raw_json"])["sku"] == "=bad"
    again = read_source("review_rows.csv", files["review_rows.csv"])
    assert len(again.rows) == 3
    with zipfile.ZipFile(io.BytesIO(files["dataforge-results.zip"])) as archive:
        assert set(archive.namelist()) == set(files) - {"dataforge-results.zip"}
    book = load_workbook(io.BytesIO(files["summary.xlsx"]), data_only=False)
    assert "normalized" not in book.sheetnames
    assert book["review_rows"]["D2"].value == "000123"
    assert book["review_rows"]["D3"].value == "=bad"
    assert book["review_rows"]["D3"].data_type == "s"
    assert book["review_rows"].max_row == 4
    book.close()


def test_xlsx_numeric_sku_and_multi_sheet():
    a = read_source("numeric_sku.xlsx", (DATA / "numeric_sku.xlsx").read_bytes())
    assert a.rows[0][0] == "000123"
    assert run(a)["counts"]["typed_sku"] == 1
    book = Workbook()
    book.create_sheet("extra")
    data = io.BytesIO()
    book.save(data)
    with pytest.raises(DataError, match="один аркуш"):
        read_source("many.xlsx", data.getvalue())


def test_xlsx_archive_expansion_limit():
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("bomb.xml", b"0" * (16 * 1024 * 1024 + 1))
    with pytest.raises(DataError, match="16 МБ"):
        read_source("bomb.xlsx", stream.getvalue())


def test_demo_expected():
    a = read_source("supplier_a.csv", (DATA / "supplier_a.csv").read_bytes())
    b = read_source("supplier_b.xlsx", (DATA / "supplier_b.xlsx").read_bytes())
    result = run(a, b)
    assert len(result["products"]) == 22
    assert result["errors"] == 12
    assert result["warnings"] == 4
    assert len(result["transforms"]) == 2
    assert len(result["changes"]) == 16
    assert not any(c["sku"] == "BAG-03" and c["kind"] == "price" for c in result["changes"])
