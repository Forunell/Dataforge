"""Reproducible fictional fixtures. Run from the repository root."""
import csv
import io
import json
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1] / "TEST_DATA"
HEADER = ["Артикул", "Назва", "Ціна", "Валюта", "Залишок"]
A = [
    ["000123", 'Чашка «Ранок»', "129,90", "UAH", "12"],
    ["LAMP-02", "Лампа Листок", "890", "UAH", "5"],
    ["BAG-03", "Торба Місто", "12.50", "USD", "8"],
    ["OLD-04", "Блокнот Пісок", "75", "UAH", "3"],
    ["DUP-05", "Свічка Тиша", "180", "UAH", "4"],
    ["DUP-05", "Свічка Тиша", "185", "UAH", "6"],
    ["NEG-06", "Підставка Хвиля", "-50", "UAH", "-2"],
    ["", "Кошик Лоза", "300", "UAH", "2"],
    ["  space-07  ", " Плед Хмара ", "450,00", "UAH", "7"],
    ["QUOTE-08", 'Набір "Лінія"\nдва предмети', "210", "UAH", "4"],
    ["TEXT-09", '=HYPERLINK("https://example.invalid")', "10", "UAH", "1"],
    ["EMPTY-10", "", "", "UAH", ""],
]
B = [
    ["000123", 'Чашка «Ранок»', "139.90", "UAH", "10"],
    ["LAMP-02", "Лампа Листок", "890", "UAH", "5"],
    ["BAG-03", "Торба Місто", "11", "EUR", "6"],
    ["NEW-11", "Ваза Обрій", "520", "UAH", "9"],
    ["SPACE-07", "Плед Хмара", "460", "UAH", "5"],
    ["QUOTE-08", 'Набір "Лінія"\nдва предмети', "210", "UAH", "4"],
    ["BAD-12", "Годинник Коло", "сто", "UAH", "1.5"],
    ["CUR-13", "Органайзер Сад", "220", "грн", "3"],
    ["NAME-14", "", "99", "UAH", "0"],
    ["=SKU-15", "Постер Тераса", "125", "UAH", "2"],
]


def write_csv(name, rows, delimiter=";", encoding="utf-8-sig", header=HEADER):
    with (ROOT / name).open("w", encoding=encoding, newline="") as stream:
        writer = csv.writer(stream, delimiter=delimiter)
        writer.writerow(header)
        writer.writerows(rows)


def main():
    ROOT.mkdir(exist_ok=True)
    write_csv("supplier_a.csv", A)
    write_csv("supplier_b.csv", B)
    clean_a = [["0001", "Чай", "10", "UAH", "2"], ["0002", "Кава", "20", "UAH", "1"], ["OLD", "Сік", "5", "UAH", "3"]]
    clean_b = [["0001", "Чай", "10.00", "UAH", "2"], ["0002", "Кава", "22", "UAH", "2"], ["NEW", "Вода", "5", "UAH", "3"]]
    write_csv("clean_a.csv", clean_a)
    write_csv("clean_b.csv", clean_b)
    (ROOT / "golden_results.json").write_text(json.dumps({
        "demo": {"rows": 22, "errors": 12, "warnings": 4, "quality": 45.5,
                 "row_counts": {"ok": 10, "warning": 4, "error": 8},
                 "comparison_counts": {"changed": 1, "unchanged": 2, "blocked": 8, "removed": 3, "new": 3},
                 "transforms": 2, "change_events": 16},
        "clean": {"rows": 6, "errors": 0, "warnings": 0, "quality": 100,
                  "comparison_counts": {"unchanged": 1, "changed": 1, "removed": 1, "new": 1}}
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv("cp1251.csv", [["000001", 'Чашка "Лист"', "12,5", "UAH", "2"]], encoding="cp1251")
    write_csv("comma.csv", [["000123", 'Текст, "лапки"\nновий рядок', "12.50", "UAH", "2"]], delimiter=",")
    write_csv("blank_row.csv", [["000001", "Чашка", "10", "UAH", "1"], ["", "", "", "", ""]])
    write_csv("injection.csv", [[prefix + "1+1", prefix + "cmd", "10", "UAH", "2"] for prefix in ["=", "+", "-", "@", "\t=", "  ="]])
    (ROOT / "broken.csv").write_text('sku;name;price;currency;stock\n1;"missing quote;1;UAH;2', encoding="utf-8")
    (ROOT / "broken.xlsx").write_bytes(b"This is not an XLSX file")
    (ROOT / "empty.csv").write_bytes(b"")
    (ROOT / "oversize.csv").write_bytes(b"x" * (2 * 1024 * 1024 + 1))
    write_csv("too_many_rows.csv", [[str(i), "Товар", "1", "UAH", "1"] for i in range(2001)])
    book = Workbook()
    sheet = book.active
    sheet.title = "Прайс B"
    sheet.append(HEADER)
    for row in B:
        sheet.append(row)
        for cell in sheet[sheet.max_row]:
            cell.data_type = "s"
            cell.number_format = "@"
    sheet.freeze_panes = "A2"
    for column in "ABCDE":
        sheet.column_dimensions[column].width = 27
    book.save(ROOT / "supplier_b.xlsx")
    typed = Workbook()
    typed.active.append(HEADER)
    typed.active.append([123, "Чашка", 10, "UAH", 2])
    typed.active["A2"].number_format = "000000"
    typed.save(ROOT / "numeric_sku.xlsx")
    stream = io.StringIO()
    writer = csv.writer(stream)
    writer.writerow(["fixture", "expected"])
    for name, expected in [
        ("supplier_a.csv", "12 rows; SKU 000123 preserved; duplicates DUP-05; negative price/stock; required fields; formula warning"),
        ("supplier_b.xlsx", "10 rows; numeric errors BAD-12; unsupported currency CUR-13; missing name NAME-14; formula SKU warning"),
        ("demo, no rules, comparison enabled", "22 rows; 12 errors; 4 warnings; 2 transformations; 16 events; 17 comparison records; quality 45.5%"),
        ("BAG-03 comparison", "currency event USD -> EUR; stock 8 -> 6; NO price event"),
        ("000123 comparison", "price 129.90 -> 139.90; stock 12 -> 10"),
        ("cp1251.csv", "auto rejected; manual cp1251 gives 1 row and Cyrillic intact"),
        ("comma.csv", "1 row; quoted comma, quotes and newline intact"),
        ("blank_row.csv", "2 rows, blank product retained with 5 required issues"),
        ("injection.csv", "6 rows; CSV dangerous text starts with protective apostrophe; XLSX string cells"),
        ("broken.csv / broken.xlsx / empty.csv", "controlled DataError; no crash or partial replacement"),
        ("oversize.csv / too_many_rows.csv", "controlled size / row limit rejection"),
        ("numeric_sku.xlsx", "000123 restored from zero format; typed_sku warning"),
    ]:
        writer.writerow([name, expected])
    (ROOT / "expected_results.csv").write_text(stream.getvalue(), encoding="utf-8-sig")
    print(f"Created fixtures in {ROOT}")


if __name__ == "__main__":
    main()
