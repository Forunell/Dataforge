"""Real Edge E2E; starts and stops an owned local server when necessary."""
import csv
import io
import json
import subprocess
import sys
import time
import urllib.request
import zipfile
from pathlib import Path

from openpyxl import Workbook
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "artifacts"
EDGE = Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
BASE = "http://127.0.0.1:8765"


def wait_heading(page, text):
    page.get_by_role("heading", name=text, exact=True).wait_for()


def screenshot(page, name, width, results):
    page.screenshot(path=str(OUT / f"{name}-{width}.png"), full_page=True)
    assert not page.evaluate("document.documentElement.scrollWidth > window.innerWidth"), f"Overflow: {name} / {width}"
    results.append(dict(width=width, scenario=name, status="PASS", horizontal_overflow=False))


def export(page, width, clean=False):
    page.locator('input[name="acknowledge"]').check()
    page.get_by_role("button", name="Підтвердити й створити експорт", exact=False).click()
    wait_heading(page, "Звіти готові до завантаження.")
    with page.expect_download() as event:
        page.get_by_role("link", name="Завантажити все", exact=False).click()
    download = event.value
    path = OUT / f"browser-{'clean' if clean else 'review'}-{width}.zip"
    download.save_as(path)
    with zipfile.ZipFile(path) as archive:
        assert ("normalized.csv" in archive.namelist()) == clean
        assert ("review_rows.csv" in archive.namelist()) != clean
        report = json.loads(archive.read("report.json"))
        assert report["status"] == ("checked" if clean else "needs_review")
        rows = list(csv.DictReader(io.StringIO(archive.read("normalized.csv" if clean else "review_rows.csv").decode("utf-8-sig")), delimiter=";"))
        assert len(rows) == (6 if clean else 22)
        if clean:
            assert rows[0]["sku"] == "0001"


def clean_payload(which):
    rows = "0001;Чай;10;UAH;2\n0002;Кава;20;UAH;1\nOLD;Сік;5;UAH;3" if which == "A" else "0001;Чай;10.00;UAH;2\n0002;Кава;22;UAH;2\nNEW;Вода;5;UAH;3"
    return dict(name=f"clean_{which}.csv", mimeType="text/csv", buffer=("sku;name;price;currency;stock\n" + rows).encode("utf-8-sig"))


def main():
    OUT.mkdir(exist_ok=True)
    owned = None
    server_log = (OUT / "browser-server.log").open("w", encoding="utf-8")
    try:
        try:
            urllib.request.urlopen(BASE + "/health", timeout=1).close()
        except OSError:
            owned = subprocess.Popen([sys.executable, "-m", "uvicorn", "dataforge.app:app", "--host", "127.0.0.1", "--port", "8765"], cwd=ROOT, stdout=server_log, stderr=server_log, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            for _ in range(50):
                try:
                    urllib.request.urlopen(BASE + "/health", timeout=1).close()
                    break
                except OSError:
                    time.sleep(.1)
            else:
                raise RuntimeError("Local server did not become ready")
        results, errors, remote, http_errors = [], [], [], []
        with sync_playwright() as p:
            browser = p.chromium.launch(executable_path=str(EDGE), headless=True)
            for width in [1440, 768, 390]:
                context = browser.new_context(viewport={"width": width, "height": 1000}, accept_downloads=True)
                page = context.new_page()
                page.on("pageerror", lambda e: errors.append(str(e)))
                page.on("request", lambda req: remote.append(req.url) if not req.url.startswith(BASE) else None)
                page.on("response", lambda res: http_errors.append([res.status, res.url]) if res.status >= 500 else None)
                page.goto(BASE)
                wait_heading(page, "Чисті дані починаються тут.")
                assert page.locator(".upload-metrics .metric strong").all_text_contents() == ["0", "0", "—", "—"]
                assert page.locator('a[href="/step/3"]').count() == 0
                screenshot(page, "step-1", width, results)
                page.get_by_role("button", name="Відкрити демо", exact=False).click()
                wait_heading(page, "Знайдемо спільну мову.")
                assert page.locator(".demo-banner").is_visible()
                assert page.locator('select[name="a_sku"]').input_value() == "0"
                page.locator('input[name="compare"]').check()
                screenshot(page, "step-2", width, results)
                page.get_by_role("button", name="Перевірити дані", exact=False).click()
                wait_heading(page, "Усі зміни — перед очима.")
                assert page.locator("[data-record]").count() == 17
                assert page.locator("[data-issue]").count() == 16
                page.locator("#issue-filter").select_option("negative")
                page.locator("#issue-filter").locator("xpath=ancestor::form").get_by_role("button", name="Показати", exact=True).click()
                assert page.locator("[data-issue]").count() == 2
                page.goto(BASE + "/step/3")
                screenshot(page, "step-3", width, results)
                # Preview, apply, undo: immutable original and current error count.
                edit = page.locator('form[action="/edit/preview"]')
                edit.locator('[name="target"]').select_option("0:2")
                edit.locator('[name="row"]').fill("8")
                edit.locator('[name="value"]').fill("50")
                edit.get_by_role("button", name="Переглянути правки").click()
                assert page.locator(".metric.danger strong").inner_text().startswith("12")
                page.get_by_role("button", name="Застосувати 1 правок").click()
                assert page.locator(".metric.danger strong").inner_text().startswith("11")
                page.get_by_role("button", name="Скасувати останню правку").click()
                assert page.locator(".metric.danger strong").inner_text().startswith("12")
                results.append(dict(width=width, scenario="manual preview/apply/undo", status="PASS"))
                export(page, width)
                assert page.locator('a[href="/download/normalized.csv"]').count() == 0
                screenshot(page, "step-4", width, results)
                page.get_by_role("button", name="Почати заново", exact=False).click()
                wait_heading(page, "Чисті дані починаються тут.")
                # Real upload and clean export.
                page.locator("#file_a").set_input_files(clean_payload("A"))
                page.locator("#file_b").set_input_files(clean_payload("B"))
                page.get_by_role("button", name="Прочитати файли", exact=False).click()
                wait_heading(page, "Знайдемо спільну мову.")
                assert page.locator(".demo-banner").count() == 0
                page.locator('input[name="compare"]').check()
                page.get_by_role("button", name="Перевірити дані", exact=False).click()
                wait_heading(page, "Усі зміни — перед очима.")
                assert page.locator("[data-record]").count() == 4
                page.locator('select[name="kind"]').select_option("unchanged")
                page.locator('select[name="kind"]').locator("xpath=ancestor::form").get_by_role("button", name="Застосувати", exact=True).click()
                assert page.locator("[data-record]").count() == 1
                assert "0001" in page.locator("[data-record]").inner_text()
                page.goto(BASE + "/step/3")
                screenshot(page, "clean-review", width, results)
                export(page, width, clean=True)
                screenshot(page, "clean-export", width, results)
                # Replacing only A invalidates the previous export and retains B.
                page.goto(BASE + "/step/1")
                page.locator("#file_a").set_input_files(clean_payload("A"))
                page.get_by_role("button", name="Прочитати файли", exact=False).click()
                wait_heading(page, "Знайдемо спільну мову.")
                assert "clean_B.csv" in page.content()
                assert context.request.get(BASE + "/download/normalized.csv").status == 404
                results.append(dict(width=width, scenario="single-file replacement invalidates export", status="PASS"))
                page.get_by_role("button", name="Почати заново", exact=False).click()
                # Encoding failure and recovery using real fixture files.
                page.locator("#file_a").set_input_files(str(ROOT / "TEST_DATA/cp1251.csv"))
                page.locator("#file_b").set_input_files(str(ROOT / "TEST_DATA/supplier_b.xlsx"))
                page.get_by_role("button", name="Прочитати файли", exact=False).click()
                page.get_by_role("alert").wait_for()
                assert "Windows-1251" in page.get_by_role("alert").inner_text()
                page.locator("#file_a").set_input_files(str(ROOT / "TEST_DATA/cp1251.csv"))
                page.locator("#file_b").set_input_files(str(ROOT / "TEST_DATA/supplier_b.xlsx"))
                page.locator("details.import-settings").first.locator("summary").click()
                page.locator('select[name="encoding_a"]').select_option("cp1251")
                page.get_by_role("button", name="Прочитати файли", exact=False).click()
                wait_heading(page, "Знайдемо спільну мову.")
                assert "000001" in page.content()
                results.append(dict(width=width, scenario="CP1251 failure and recovery", status="PASS"))
                if width == 1440:
                    page.get_by_role("button", name="Почати заново", exact=False).click()
                    book = Workbook()
                    book.active.title = "Empty"
                    sheet = book.create_sheet("Prices")
                    sheet.append(["sku", "name", "price", "currency", "stock"])
                    sheet.append(["0001", "Чай", 10, "UAH", 1])
                    buf = io.BytesIO()
                    book.save(buf)
                    page.locator("#file_a").set_input_files(dict(name="multi.xlsx", mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer=buf.getvalue()))
                    # Drag-and-drop through the same event path as a real browser drop.
                    payload = clean_payload("B")["buffer"].decode("utf-8-sig")
                    page.locator('label[for="file_b"]').evaluate("(el, text) => { const transfer = new DataTransfer(); transfer.items.add(new File([text], 'dragged.csv', {type:'text/csv'})); el.dispatchEvent(new DragEvent('drop', {dataTransfer:transfer,bubbles:true})); }", payload)
                    assert page.locator("#file_b").evaluate("el => el.files[0].name") == "dragged.csv"
                    page.get_by_role("button", name="Прочитати файли", exact=False).click()
                    wait_heading(page, "Виберіть аркуші Excel")
                    page.locator('select[name="sheet_0"]').select_option("Prices")
                    page.get_by_role("button", name="Прочитати вибрані аркуші", exact=False).click()
                    wait_heading(page, "Знайдемо спільну мову.")
                    assert "Prices" in page.content()
                    results.append(dict(width=width, scenario="drag/drop and multisheet with empty first sheet", status="PASS"))
                page.get_by_role("button", name="Почати заново", exact=False).click()
                context.close()
            browser.close()
        assert not errors, errors
        assert not remote, remote
        assert not http_errors, http_errors
        report = dict(date="2026-11-07", engine="Microsoft Edge / Playwright", results=results, js_errors=errors, external_app_requests=remote, server_errors=http_errors)
        (OUT / "browser-results.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(report, ensure_ascii=True, indent=2))
    finally:
        if owned:
            owned.terminate()
            owned.wait(timeout=10)
        server_log.close()


if __name__ == "__main__":
    main()
