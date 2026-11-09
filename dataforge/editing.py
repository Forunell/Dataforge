"""Reversible cell patches; original sources never change."""
from copy import deepcopy
from decimal import InvalidOperation
import json

from dataforge.core import DataError, MAX_CELL, number


def preview_edits(sources, slot, column, operation, row=None, value=""):
    if slot not in (0, 1) or not 0 <= column < len(sources[slot].headers):
        raise DataError("Виберіть джерело та колонку.")
    source = sources[slot]
    if operation not in {"manual", "trim", "upper", "lower", "decimal"}:
        raise DataError("Невідома операція.")
    if operation == "manual" and (row is None or not 2 <= row <= len(source.rows) + 1):
        raise DataError("Вкажіть номер рядка з попереднього перегляду.")
    changes = []
    for index, cells in enumerate(source.rows, 2):
        if operation == "manual" and index != row:
            continue
        before = cells[column]
        if operation == "manual":
            after = value
        elif operation == "trim":
            after = before.strip()
        elif operation == "upper":
            after = before.upper()
        elif operation == "lower":
            after = before.lower()
        else:
            try:
                after = format(number(before), "f")
            except (InvalidOperation, DataError):
                continue  # Ambiguous and invalid values remain untouched and in the quality report.
        if len(after) > MAX_CELL or any(ord(c) < 32 and c not in "\t\r\n" for c in after):
            raise DataError("Значення перевищує 1000 символів або містить недопустимі керівні символи.")
        if before != after:
            changes.append(dict(slot=slot, source="AB"[slot], row=index, ci=column,
                                column=source.headers[column], before=before, after=after))
    return changes


def apply_edits(sources, patches, reverse=False):
    updated = deepcopy(sources)
    for p in patches:
        cells = updated[p["slot"]].rows[p["row"] - 2]
        expected, value = (p["after"], p["before"]) if reverse else (p["before"], p["after"])
        if cells[p["ci"]] != expected:
            raise DataError("Попередній перегляд застарів. Повторіть дію.")
        cells[p["ci"]] = value
        source = updated[p["slot"]]
        if len(json.dumps(dict(zip(source.headers, cells, strict=True)), ensure_ascii=False)) > 30000:
            raise DataError("Рядок після правки перевищує 30000 символів для Excel-звіту.")
    return updated
