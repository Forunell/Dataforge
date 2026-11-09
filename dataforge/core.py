import csv
import io
import json
import re
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook

MAX_BYTES = 2 * 1024 * 1024
MAX_ROWS = 2000
MAX_COLS = 64
MAX_CELL = 1000
FIELDS = {"sku": "SKU", "name": "Назва", "price": "Ціна", "currency": "Валюта", "stock": "Залишок",
          "category": "Категорія", "variant": "Варіант"}
CURRENCIES = {"UAH", "USD", "EUR", "GBP", "PLN", "CZK", "CHF", "CAD", "AUD", "JPY", "CNY"}
ALIASES = {
    "sku": {"sku", "артикул", "код", "код товару"},
    "name": {"name", "назва", "название", "товар"},
    "price": {"price", "ціна", "цена"},
    "currency": {"currency", "валюта"},
    "stock": {"stock", "залишок", "остаток"},
    "category": {"category", "категорія", "категория"},
    "variant": {"variant", "варіант", "вариант", "розмір", "колір"},
}
LABELS = {
    "required": "Обов’язкове поле", "number": "Некоректне число", "negative": "Від’ємне значення",
    "integer": "Дробовий залишок", "currency": "Непідтримувана валюта",
    "duplicate": "Дублікат у файлі", "shared_sku": "SKU в обох файлах",
    "collision": "Колізія після зміни SKU", "formula": "Формульний текст",
    "typed_sku": "SKU не був текстом в Excel", "identity": "Різні назви для одного SKU",
    "currency_mismatch": "Різні або невідомі валюти", "threshold": "Поріг зміни ціни",
    "ambiguous": "Неоднозначне число", "key": "Неповний ключ зіставлення",
}


class DataError(ValueError):
    pass


@dataclass
class Source:
    name: str
    format: str
    headers: list[str]
    rows: list[list[str]]
    encoding: str = "—"
    delimiter: str = "—"
    typed: set[tuple[int, int]] = field(default_factory=set)
    size: int = 0
    sheets: list[str] = field(default_factory=list)
    sheet: str = ""


def workbook_sheets(data):
    if len(data) > MAX_BYTES:
        raise DataError("Файл перевищує 2 МБ.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if len(archive.infolist()) > 1000 or sum(i.file_size for i in archive.infolist()) > 16 * 1024 * 1024:
                raise DataError("XLSX перевищує ліміт розпакованих даних (16 МБ).")
        book = load_workbook(io.BytesIO(data), read_only=True, data_only=False, keep_links=False)
        try:
            return book.sheetnames
        finally:
            book.close()
    except DataError:
        raise
    except Exception as exc:
        raise DataError("Не вдалося відкрити XLSX. Перевірте формат книги.") from exc


def read_source(name, data, encoding="auto", delimiter="auto", sheet_name=None):
    if len(data) > MAX_BYTES:
        raise DataError("Файл перевищує 2 МБ.")
    if not data:
        raise DataError("Файл порожній.")
    suffix = Path(name).suffix.lower()
    typed = set()
    sheets, selected_sheet = [], ""
    if suffix == ".csv":
        if encoding not in {"auto", "utf-8-sig", "cp1251"} or delimiter not in {"auto", ",", ";", "\t"}:
            raise DataError("Неприпустиме кодування або роздільник.")
        try:
            # Legacy encoding is deliberately never guessed.
            actual_encoding = "utf-8-sig" if encoding == "auto" else encoding
            content = data.decode(actual_encoding)
        except UnicodeError as exc:
            raise DataError("Не вдалося прочитати UTF-8. Виберіть Windows-1251 вручну.") from exc
        if "\x00" in content:
            raise DataError("CSV містить нульові байти; перевірте формат.")
        if delimiter == "auto":
            try:
                delimiter = csv.Sniffer().sniff(content[:65536], delimiters=",;\t").delimiter
            except csv.Error as exc:
                raise DataError("Роздільник неоднозначний. Виберіть його вручну.") from exc
        try:
            rows = []
            for row in csv.reader(io.StringIO(content, newline=""), delimiter=delimiter, strict=True):
                rows.append(row)
                if len(rows) > MAX_ROWS + 1:
                    raise DataError(f"Максимум {MAX_ROWS} рядків даних на файл.")
        except csv.Error as exc:
            raise DataError("Пошкоджений CSV: перевірте лапки та довжину полів.") from exc
        encoding = actual_encoding
    elif suffix == ".xlsx":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                infos = archive.infolist()
                if len(infos) > 1000 or sum(i.file_size for i in infos) > 16 * 1024 * 1024:
                    raise DataError("XLSX перевищує ліміт розпакованих даних (16 МБ).")
            book = load_workbook(io.BytesIO(data), read_only=True, data_only=False, keep_links=False)
            try:
                sheets = book.sheetnames
                if len(sheets) > 1 and sheet_name is None:
                    raise DataError("Виберіть один аркуш XLSX для імпорту.")
                if sheet_name is not None and sheet_name not in sheets:
                    raise DataError("Вибраний аркуш відсутній у книзі.")
                sheet = book[sheet_name] if sheet_name else book.worksheets[0]
                selected_sheet = sheet.title
                if (sheet.max_row or 0) > MAX_ROWS + 1 or (sheet.max_column or 0) > MAX_COLS:
                    raise DataError("XLSX перевищує ліміт 2000 рядків або 64 колонок.")
                sheet.reset_dimensions()
                rows = []
                for ri, cells in enumerate(sheet.iter_rows()):
                    if ri > MAX_ROWS or len(cells) > MAX_COLS:
                        raise DataError("XLSX перевищує ліміт рядків або колонок.")
                    row = []
                    for ci, cell in enumerate(cells):
                        value = cell.value
                        if value is None:
                            text = ""
                        elif isinstance(value, (date, datetime)):
                            text = value.isoformat()
                        elif isinstance(value, int) and re.fullmatch(r"0{1,64}", cell.number_format or ""):
                            text = str(value).zfill(len(cell.number_format))
                        else:
                            text = str(value)
                        if value is not None and not isinstance(value, str):
                            typed.add((ri + 1, ci))
                        row.append(text)
                    rows.append(row)
            finally:
                book.close()
        except DataError:
            raise
        except Exception as exc:
            raise DataError("Не вдалося відкрити XLSX. Потрібна незашифрована книга Excel.") from exc
        encoding, delimiter = "—", "—"
    else:
        raise DataError("Підтримуються лише CSV та XLSX.")
    if not rows or not rows[0] or len(rows) == 1:
        raise DataError("Потрібні заголовки та хоча б один рядок даних.")
    headers = rows[0]
    if len(headers) > MAX_COLS or any(not h.strip() for h in headers):
        raise DataError("Потрібні непорожні заголовки; максимум 64 колонки.")
    if len(set(headers)) != len(headers):
        raise DataError("Заголовки повторюються. Надайте колонкам різні назви.")
    for ri, row in enumerate(rows):
        # Preserve blank records as blank products, never silently drop them.
        if not row:
            rows[ri] = [""] * len(headers)
        elif suffix == ".xlsx" and len(row) < len(headers):
            rows[ri] = row + [""] * (len(headers) - len(row))
        elif len(row) != len(headers):
            raise DataError(f"Рядок {ri + 1}: кількість колонок не збігається із заголовками.")
        if any(len(v) > MAX_CELL for v in rows[ri]):
            raise DataError(f"Рядок {ri + 1}: поле перевищує {MAX_CELL} символів.")
        if any(ord(c) < 32 and c not in "\t\r\n" for v in rows[ri] for c in v):
            raise DataError(f"Рядок {ri + 1}: недопустимі керівні символи.")
        if len(json.dumps(dict(zip(headers, rows[ri], strict=True)), ensure_ascii=False)) > 30000:
            raise DataError(f"Рядок {ri + 1}: початкові значення перевищують ліміт 30000 символів для Excel-звіту.")
    return Source(Path(name.replace("\\", "/")).name[:200], suffix[1:].upper(), headers, rows[1:],
                  encoding, delimiter, typed, len(data), sheets, selected_sheet)


def suggest(source):
    result = {}
    for key, aliases in ALIASES.items():
        candidates = [i for i, h in enumerate(source.headers) if h.strip().lower() in aliases]
        result[key] = candidates[0] if len(candidates) == 1 else -1
    return result


def dangerous(value):
    return bool(value) and value.lstrip(" \t\r\n\ufeff").startswith(("=", "+", "-", "@"))


def safe_text(value):
    value = str(value)
    return "'" + value if dangerous(value) or re.match(r"^[ \ufeff]*[\x00-\x1f\x7f]", value) else value


def number(value):
    """Exact decimal parsing. A lone 3-digit suffix is deliberately ambiguous."""
    value = value.strip().replace("\u00a0", " ").replace("\u202f", " ")
    if len(value) > 40:
        raise InvalidOperation
    sign = ""
    if value.startswith(("+", "-")):
        sign, value = value[0], value[1:]
    if " " in value:
        if not re.fullmatch(r"[0-9]{1,3}(?: [0-9]{3})+(?:[.,][0-9]{1,6})?", value):
            raise InvalidOperation
        value = value.replace(" ", "")
    elif "," in value and "." in value:
        decimal = "," if value.rfind(",") > value.rfind(".") else "."
        grouping = "." if decimal == "," else ","
        pattern = r"[0-9]{1,3}(?:" + re.escape(grouping) + r"[0-9]{3})+" + re.escape(decimal) + r"[0-9]{1,6}"
        if not re.fullmatch(pattern, value):
            raise InvalidOperation
        value = value.replace(grouping, "")
    elif re.fullmatch(r"[0-9]{1,3}(,[0-9]{3}){2,}", value) or re.fullmatch(r"[0-9]{1,3}(\.[0-9]{3}){2,}", value):
        value = value.replace(",", "").replace(".", "")
    elif re.fullmatch(r"[0-9]{1,3}[.,][0-9]{3}", value):
        raise DataError("Неоднозначне число: 1,234 може означати 1234 або 1.234. Уточніть запис вручну.")
    if not re.fullmatch(r"[0-9]+(?:[.,][0-9]{1,6})?", value) or len(re.sub(r"[^0-9]", "", value)) > 18:
        raise InvalidOperation
    return Decimal(sign + value.replace(",", "."))


def validate_mapping(source, mapping, rules):
    if set(mapping) != set(FIELDS) or any(type(v) is not int or v < -1 or v >= len(source.headers) for v in mapping.values()):
        raise DataError("Некоректне зіставлення колонок.")
    required = {"sku", "name", "price"}
    if rules.get("composite"):
        required.add("variant")
    if any(mapping[k] < 0 for k in required):
        raise DataError("Виберіть колонки SKU, назви, ціни та варіанта для складеного ключа.")
    selected = [v for v in mapping.values() if v >= 0]
    if len(set(selected)) != len(selected):
        raise DataError("Кожне цільове поле повинно мати окрему вихідну колонку.")


def analyze(sources, mappings, rules, compare=False):
    if len(sources) != 2 or len(mappings) != 2:
        raise DataError("Потрібні два файли та два зіставлення.")
    products, issues, transforms = [], [], []
    threshold = Decimal(str(rules.get("threshold", "20")))
    if not threshold.is_finite() or not 0 <= threshold <= 100000:
        raise DataError("Поріг зміни ціни: від 0 до 100000 відсотків.")

    def issue(p, code, field_name, message, severity="error"):
        si = "AB".index(p["source"])
        ci = mappings[si].get(field_name, -1)
        raw = sources[si].rows[p["row"] - 2]
        issues.append(dict(source=p["source"], file=p["file"], sheet=sources[si].sheet,
                           row=p["row"], sku=p["sku"], code=code, field=field_name,
                           column=sources[si].headers[ci] if ci >= 0 else FIELDS.get(field_name, field_name),
                           original=raw[ci] if ci >= 0 else "", current=raw[ci] if ci >= 0 else "",
                           severity=severity, message=message))

    for si, (source, mapping) in enumerate(zip(sources, mappings, strict=True)):
        validate_mapping(source, mapping, rules)
        for ri, raw in enumerate(source.rows, 2):
            original = {k: raw[v] if v >= 0 else "" for k, v in mapping.items()}
            values = original.copy()
            for key in FIELDS:
                if key != "sku" and rules.get("trim"):
                    values[key] = values[key].strip()
            if rules.get("sku_trim"):
                values["sku"] = values["sku"].strip()
            if rules.get("sku_spaces"):
                values["sku"] = " ".join(values["sku"].split())
            if rules.get("sku_upper"):
                values["sku"] = values["sku"].upper()
            p = dict(source="AB"[si], file=source.name, sheet=source.sheet, row=ri, **values,
                     raw_sku=original["sku"], raw_json=json.dumps(dict(zip(source.headers, raw, strict=True)), ensure_ascii=False))
            p["working_json"] = p["raw_json"]
            required = {"price"}
            if rules.get("require_sku", True):
                required.add("sku")
            if rules.get("require_name", True):
                required.add("name")
            required.update(k for k in ("currency", "stock") if mapping[k] >= 0)
            if rules.get("composite"):
                required.add("variant")
            for key in FIELDS:
                if key in required and not values[key].strip():
                    issue(p, "required", key, f"Заповніть поле «{FIELDS[key]}».")
                if key not in {"price", "stock"} and dangerous(values[key]):
                    issue(p, "formula", key, "Текст схожий на формулу; CSV отримає захисний апостроф.", "warning")
            if not values["sku"].strip() and "sku" not in required:
                issue(p, "key", "sku", "Без SKU рядок збережено, але зіставлення неможливе.", "warning")
            if (ri, mapping["sku"]) in source.typed:
                issue(p, "typed_sku", "sku", "Excel зберіг SKU не як текст. Перевірте нулі та точність у першоджерелі.", "warning")
            for key in ("price", "stock"):
                if not values[key].strip():
                    continue
                try:
                    n = number(values[key])
                    p[key] = format(n, "f")
                    if n < 0 and rules.get("nonnegative_" + key, True):
                        issue(p, "negative", key, "Значення не може бути від’ємним.")
                    if key == "stock" and rules.get("integer_stock", True) and n != n.to_integral_value():
                        issue(p, "integer", key, "Залишок має бути цілим числом.")
                except DataError as exc:
                    issue(p, "ambiguous", key, str(exc))
                except InvalidOperation:
                    issue(p, "number", key, "Потрібне точне число: до 18 цифр загалом і до 6 після десяткового знака.")
            if values["currency"] and values["currency"] not in CURRENCIES:
                issue(p, "currency", "currency", "Вкажіть підтримуваний код валюти великими латинськими літерами.")
            for key in FIELDS:
                if original[key] != p[key]:
                    transforms.append(dict(source=p["source"], row=ri, field=key, before=original[key], after=p[key]))
            p["key"] = (p["sku"], p["variant"]) if rules.get("composite") else (p["sku"],)
            products.append(p)
    groups = defaultdict(list)
    missing = []
    for p in products:
        if all(part.strip() for part in p["key"]):
            groups[p["key"]].append(p)
        else:
            missing.append(p)
    for group in groups.values():
        counts = Counter(p["source"] for p in group)
        collision = len({p["raw_sku"] for p in group}) > 1
        for p in group:
            if counts[p["source"]] > 1:
                issue(p, "duplicate", "sku", "Ключ повторюється у файлі. Усі рядки збережені; зіставлення призупинено.",
                      "error" if rules.get("unique", True) else "warning")
            if collision:
                issue(p, "collision", "sku", "Різні початкові SKU стали однаковими після нормалізації.")
    invalid = {(i["source"], i["row"]) for i in issues if i["severity"] == "error"}
    changes, records = [], []

    def record(group, kind, a, b):
        p = group[0]
        return dict(sku=p["sku"], variant=p["variant"], name=(b or a or [p])[0]["name"], kind=kind,
                    values_a={k: a[0][k] for k in FIELDS} if len(a) == 1 else {},
                    values_b={k: b[0][k] for k in FIELDS} if len(b) == 1 else {},
                    rows_a=", ".join(str(x["row"]) for x in a), rows_b=", ".join(str(x["row"]) for x in b),
                    before=json.dumps({k: a[0][k] for k in FIELDS}, ensure_ascii=False) if len(a) == 1 else "",
                    after=json.dumps({k: b[0][k] for k in FIELDS}, ensure_ascii=False) if len(b) == 1 else "")

    if compare:
        for group in groups.values():
            a = [p for p in group if p["source"] == "A"]
            b = [p for p in group if p["source"] == "B"]
            base = dict(sku=group[0]["sku"], variant=group[0]["variant"], file_a=sources[0].name,
                        file_b=sources[1].name, sheet_a=sources[0].sheet, sheet_b=sources[1].sheet,
                        rows_a=", ".join(str(p["row"]) for p in a), rows_b=", ".join(str(p["row"]) for p in b))

            def change(kind, before="", after="", message="", field_name=""):
                changes.append(dict(base, kind=kind, before=before, after=after, message=message, field=field_name,
                                    column_a=sources[0].headers[mappings[0][field_name]] if field_name and mappings[0][field_name] >= 0 else "",
                                    column_b=sources[1].headers[mappings[1][field_name]] if field_name and mappings[1][field_name] >= 0 else ""))

            if any((p["source"], p["row"]) in invalid for p in group) or len(a) > 1 or len(b) > 1:
                kind = "blocked"
                change(kind, message="Порівняння призупинено: помилки або дублікати.")
            elif not a:
                kind = "new"
                change(kind, after=b[0]["name"], message="Є лише у B.")
            elif not b:
                kind = "removed"
                change(kind, before=a[0]["name"], message="Є лише у A.")
            else:
                old, new = a[0], b[0]
                before_count = len(changes)
                currency_ok = old["currency"] == new["currency"] and bool(old["currency"])
                if not currency_ok:
                    change("currency", old["currency"], new["currency"], "Порівняння ціни неможливе без курсу.", "currency")
                    for p in group:
                        issue(p, "currency_mismatch", "currency", "Валюти відрізняються або не задані. Ціни не порівнюються.", "warning")
                elif Decimal(old["price"]) != Decimal(new["price"]):
                    change("price", old["price"], new["price"], old["currency"], "price")
                    old_price, new_price = Decimal(old["price"]), Decimal(new["price"])
                    if old_price == 0 or abs(new_price - old_price) * 100 > abs(old_price) * threshold:
                        issue(new, "threshold", "price", f"Зміна ціни перевищує {threshold}% або базова ціна дорівнює нулю.", "warning")
                for key in ("name", "stock", "category", "variant"):
                    left, right = old[key], new[key]
                    if key == "name":
                        if rules.get("name_trim"):
                            left, right = " ".join(left.split()), " ".join(right.split())
                        if rules.get("name_casefold"):
                            left, right = left.casefold(), right.casefold()
                    if key == "stock" and left and right:
                        different = Decimal(left) != Decimal(right)
                    else:
                        different = left != right
                    if different:
                        change(key, old[key], new[key], f"Змінилося поле «{FIELDS[key]}».", key)
                kind = "blocked" if not currency_ok else "changed" if len(changes) > before_count else "unchanged"
            records.append(record(group, kind, a, b))
        for p in missing:
            records.append(record([p], "blocked", [p] if p["source"] == "A" else [], [p] if p["source"] == "B" else []))
    by_row = defaultdict(list)
    for item in issues:
        by_row[(item["source"], item["row"])].append(item)
    for p in products:
        row_issues = by_row[(p["source"], p["row"])]
        p["status"] = "error" if any(i["severity"] == "error" for i in row_issues) else "warning" if row_issues else "ok"
    row_counts = {k: sum(p["status"] == k for p in products) for k in ("ok", "warning", "error")}
    return dict(products=products, issues=issues, transforms=transforms, changes=changes, records=records,
                comparison_counts=dict(Counter(r["kind"] for r in records)), counts=dict(Counter(i["code"] for i in issues)),
                errors=sum(i["severity"] == "error" for i in issues), warnings=sum(i["severity"] == "warning" for i in issues),
                compare=compare, rules=rules.copy(), ok=row_counts["ok"], row_counts=row_counts,
                quality=round(100 * row_counts["ok"] / len(products), 1) if products else None,
                transformed_rows=len({(t["source"], t["row"]) for t in transforms}),
                checked_at=datetime.now().astimezone().isoformat(),
                sources=[dict(name=s.name, sheet=s.sheet, rows=len(s.rows), bytes=s.size) for s in sources])


NORMAL_COLUMNS = ["source", "file", "row", *FIELDS, "sheet", "status", "raw_sku", "raw_json", "working_json"]
ISSUE_COLUMNS = ["source", "file", "sheet", "row", "sku", "severity", "code", "field", "column", "original", "current", "message"]
CHANGE_COLUMNS = ["sku", "variant", "kind", "file_a", "file_b", "sheet_a", "sheet_b", "rows_a", "rows_b",
                  "field", "column_a", "column_b", "before", "after", "message"]
TRANSFORM_COLUMNS = ["source", "row", "field", "before", "after"]


def csv_bytes(rows, columns, delimiter=";"):
    stream = io.StringIO(newline="")
    writer = csv.writer(stream, delimiter=delimiter, quoting=csv.QUOTE_ALL, lineterminator="\r\n")
    writer.writerow(columns)
    for row in rows:
        writer.writerow([str(row.get(c, "")) if c in {"price", "stock"} and
                         re.fullmatch(r"-?[0-9]+(?:\.[0-9]+)?", str(row.get(c, "")))
                         else safe_text(row.get(c, "")) for c in columns])
    return stream.getvalue().encode("utf-8-sig")


def export_files(result, delimiter=";"):
    if delimiter not in {";", ","}:
        raise DataError("Неприпустимий роздільник експорту.")
    blocked = result["errors"] > 0 or result["comparison_counts"].get("blocked", 0) > 0
    tables = [("review_rows" if blocked else "normalized", result["products"], NORMAL_COLUMNS), ("issues", result["issues"], ISSUE_COLUMNS),
              ("changes", result["changes"], CHANGE_COLUMNS), ("transformations", result["transforms"], TRANSFORM_COLUMNS)]
    tables.append(("comparison", result["records"], ["sku", "variant", "name", "kind", "rows_a", "rows_b", "before", "after"]))
    tables.append(("edits", result.get("edits", []), ["source", "row", "column", "before", "after"]))
    files = {name + ".csv": csv_bytes(rows, cols, delimiter) for name, rows, cols in tables}
    book = Workbook()
    book.remove(book.active)
    for name, rows, columns in tables:
        sheet = book.create_sheet(name)
        sheet.append(columns)
        for row_number, row in enumerate(rows, 2):
            # Explicit strings keep zeros and prevent formula execution in XLSX.
            sheet.append([str(row.get(c, "")) for c in columns])
            for column_number in range(1, len(columns) + 1):
                cell = sheet.cell(row=row_number, column=column_number)
                cell.data_type = "s"
                cell.number_format = "@"
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = 24
    stream = io.BytesIO()
    book.save(stream)
    files["summary.xlsx"] = stream.getvalue()
    summary = {key: result.get(key) for key in ["checked_at", "sources", "errors", "warnings", "row_counts", "quality",
                                              "comparison_counts", "rules", "compare", "initial_errors", "resolved_errors", "new_errors"]}
    summary.update(status="needs_review" if blocked else "checked", edited_cells=len(result.get("edits", [])))
    files["report.json"] = json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8")
    files["READ_ME.txt"] = (("Основний експорт заблоковано. review_rows — робоча копія для перевірки, НЕ очищений каталог.\n" if blocked
                             else "Перевірений каталог: normalized.\n") +
                            "DataForge: A -> B. Усі рядки обох джерел збережені; status=error потребує виправлення.\n"
                            "CSV: UTF-8-SIG, усі поля в лапках. Імпортуйте SKU як текст або використайте summary.xlsx.\n"
                            "Формульний текст у CSV має захисний апостроф. raw_json містить початкові значення.\n"
                            "Валюти не конвертуються. Правила: " + json.dumps(result["rules"], ensure_ascii=False) +
                            f"\nПорівняння ввімкнене: {result['compare']}\n").encode("utf-8-sig")
    archive = io.BytesIO()
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    files["dataforge-results.zip"] = archive.getvalue()
    return files
