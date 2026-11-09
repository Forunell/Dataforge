import secrets
import time
import json
from copy import deepcopy
from decimal import InvalidOperation
from pathlib import Path
from urllib.parse import urlencode

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.concurrency import run_in_threadpool
from starlette.middleware.trustedhost import TrustedHostMiddleware
from starlette.formparsers import MultiPartParser

from dataforge.core import FIELDS, LABELS, MAX_BYTES, DataError, analyze, export_files, read_source, suggest, workbook_sheets
from dataforge.editing import apply_edits, preview_edits

ROOT = Path(__file__).resolve().parent.parent
# The middleware caps the entire request below this size; accepted uploads stay in RAM.
MultiPartParser.spool_max_size = 2 * MAX_BYTES + 65536
app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
app.add_middleware(TrustedHostMiddleware, allowed_hosts=["127.0.0.1", "localhost", "testserver"])
app.mount("/static", StaticFiles(directory=ROOT / "static"), name="static")
templates = Jinja2Templates(directory=ROOT / "templates")
SESSIONS = {}
TTL = 3600
MAX_SESSIONS = 6
RULES = ["trim", "sku_trim", "sku_spaces", "sku_upper", "composite", "name_trim", "name_casefold"]
VALIDATION_RULES = ["require_sku", "require_name", "unique", "nonnegative_price", "nonnegative_stock", "integer_stock"]
KINDS = {"all": "Усі", "new": "Нові", "removed": "Відсутні у B", "changed": "Змінені",
         "unchanged": "Без змін", "blocked": "Конфлікти", "error": "Помилки", "warning": "Попередження"}


@app.middleware("http")
async def local_guard(request: Request, call_next):
    if request.method == "POST":
        origin = request.headers.get("origin")
        if origin and origin != str(request.base_url).rstrip("/"):
            return Response("Сторонній запит відхилено.", status_code=403)
        body = bytearray()
        async for part in request.stream():
            body.extend(part)
            if len(body) > 2 * MAX_BYTES + 65536:
                return HTMLResponse("Запит завеликий. Максимум 2 МБ на файл. Поверніться назад.", status_code=413)
        request._body = bytes(body)
    response = await call_next(request)
    response.headers["Content-Security-Policy"] = "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "same-origin"
    response.headers["Cache-Control"] = "no-store"
    return response


def session(request):
    now = time.monotonic()
    for key in list(SESSIONS):
        if now - SESSIONS[key]["touched"] > TTL:
            SESSIONS.pop(key, None)
    sid = request.cookies.get("df_session")
    if sid not in SESSIONS:
        if len(SESSIONS) >= MAX_SESSIONS:
            # Do not silently evict somebody else's work.
            raise DataError("Відкрито 6 сесій. Закрийте роботу через «Почати заново» або зачекайте годину.")
        sid = secrets.token_urlsafe(32)
        SESSIONS[sid] = {"csrf": secrets.token_urlsafe(32), "touched": now, "sources": [], "result": None, "files": None}
    state = SESSIONS[sid]
    state["touched"] = now
    return sid, state


def render(request, sid, state, step=1, error=None, status=200):
    view = result_view(request, state) if state.get("result") else {}
    response = templates.TemplateResponse(request=request, name="workspace.html", context={
        "state": state, "step": step, "error": error, "fields": FIELDS, "labels": LABELS,
        "view": view, "kinds": KINDS, "validation_rules": VALIDATION_RULES,
    }, status_code=status)
    response.set_cookie("df_session", sid, httponly=True, samesite="strict", max_age=TTL)
    return response


def result_view(request, state):
    result = state["result"]
    query = request.query_params
    search, kind = query.get("q", "")[:200].casefold(), query.get("kind", "all")
    rows = result["records"]
    if kind in {"error", "warning"} or not result["compare"]:
        rows = [dict(p, kind=p["status"], rows_a=str(p["row"]) if p["source"] == "A" else "",
                     rows_b=str(p["row"]) if p["source"] == "B" else "", before=p["raw_json"], after="")
                for p in result["products"] if kind not in {"error", "warning"} or p["status"] == kind]
    elif kind != "all":
        rows = [r for r in rows if r["kind"] == kind]
    rows = [r for r in rows if search in (r["sku"] + " " + r["name"]).casefold()]
    sort = query.get("sort", "sku")
    rows = sorted(rows, key=lambda r: r.get(sort if sort in {"sku", "name", "kind"} else "sku", "").casefold(),
                  reverse=query.get("direction") == "desc")
    try:
        page = max(1, min(int(query.get("page", "1")), max(1, (len(rows) + 49) // 50)))
        issue_page = max(1, int(query.get("issue_page", "1")))
    except ValueError:
        page, issue_page = 1, 1
    code = query.get("issue", "all")
    issues = [i for i in result["issues"] if code == "all" or i["code"] == code]
    issue_pages = max(1, (len(issues) + 49) // 50)
    issue_page = min(issue_page, issue_pages)

    def link(**changes):
        return "/step/3?" + urlencode({**dict(query), **changes})

    return dict(rows=rows[(page-1)*50:page*50], total=len(rows), page=page, pages=max(1, (len(rows)+49)//50),
                q=query.get("q", ""), kind=kind, sort=sort, direction=query.get("direction", "asc"),
                issues=issues[(issue_page-1)*50:issue_page*50], issue=code, issue_page=issue_page,
                issue_pages=issue_pages, issue_total=len(issues), link=link)


def set_sources(state, sources, demo=False, originals=None, demo_sources=None):
    demo_sources = demo_sources or [demo, demo]
    state.update(sources=sources, originals=deepcopy(originals or sources), mappings=[suggest(s) for s in sources],
                 result=None, files=None, rules={}, compare=False, demo=any(demo_sources), demo_sources=demo_sources,
                 pending=None, history=[], edit_preview=None,
                 baseline=None, revision=secrets.token_urlsafe(20))


def enrich_result(state, result):
    baseline = state.get("baseline")
    errors = {(i["source"], i["row"], i["field"], i["code"]) for i in result["issues"] if i["severity"] == "error"}
    if baseline is None:
        state["baseline"] = errors
        baseline = errors
    result.update(initial_errors=len(baseline), resolved_errors=len(baseline-errors), new_errors=len(errors-baseline))
    result["edits"] = []
    for si, (original, working) in enumerate(zip(state["originals"], state["sources"], strict=True)):
        for ri, (old, new) in enumerate(zip(original.rows, working.rows, strict=True), 2):
            for ci, (before, after) in enumerate(zip(old, new, strict=True)):
                if before != after:
                    result["edits"].append(dict(source="AB"[si], row=ri, column=original.headers[ci], before=before, after=after))
    for p in result["products"]:
        si = "AB".index(p["source"])
        source = state["originals"][si]
        p["raw_json"] = json.dumps(dict(zip(source.headers, source.rows[p["row"]-2], strict=True)), ensure_ascii=False)
        p["raw_sku"] = source.rows[p["row"]-2][state["mappings"][si]["sku"]]
    for issue in result["issues"]:
        si = "AB".index(issue["source"])
        ci = state["mappings"][si].get(issue["field"], -1)
        issue["original"] = state["originals"][si].rows[issue["row"]-2][ci] if ci >= 0 else ""
    return result


def check_csrf(state, form):
    if not secrets.compare_digest(str(form.get("csrf", "")), state["csrf"]):
        raise DataError("Сесія форми застаріла. Оновіть сторінку та повторіть дію.")


@app.exception_handler(DataError)
async def data_error(request, exc):
    # Escaping remains Jinja's responsibility, including parser errors.
    return templates.TemplateResponse(request=request, name="error.html", context={"error": str(exc)}, status_code=400)


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    sid, state = session(request)
    step = 4 if state["files"] else 3 if state["result"] else 2 if state["sources"] else 1
    return render(request, sid, state, step)


@app.get("/step/{step}", response_class=HTMLResponse)
async def step_page(request: Request, step: int):
    sid, state = session(request)
    if step == 1 or (step == 2 and state["sources"]) or (step == 3 and state["result"]) or (step == 4 and state["files"]):
        return render(request, sid, state, step)
    return RedirectResponse("/", status_code=303)


@app.post("/upload")
async def upload(request: Request):
    sid, state = session(request)
    async with request.form(max_files=2, max_fields=12) as form:
        check_csrf(state, form)
        try:
            pending = []
            changed = False
            for si, slot in enumerate("ab"):
                item = form.get("file_" + slot)
                if not item or not getattr(item, "filename", None):
                    if len(state["sources"]) == 2:
                        pending.append({"source": state["sources"][si], "original": state["originals"][si],
                                        "demo": state.get("demo_sources", [False, False])[si]})
                        continue
                    raise DataError("Виберіть обидва файли постачальників.")
                changed = True
                data = await item.read(MAX_BYTES + 1)
                sheets = await run_in_threadpool(workbook_sheets, data) if item.filename.lower().endswith(".xlsx") else []
                pending.append(dict(name=item.filename, data=data, sheets=sheets,
                                    encoding=str(form.get("encoding_" + slot, "auto")),
                                    delimiter=str(form.get("delimiter_" + slot, "auto"))))
            if not changed:
                raise DataError("Виберіть хоча б один новий файл для заміни.")
            if any(len(p.get("sheets", [])) > 1 for p in pending):
                state.update(pending=pending, result=None, files=None, revision=secrets.token_urlsafe(20))
                return render(request, sid, state, 1)
            sources, originals = await parse_pending(pending, {})
        except DataError as exc:
            return render(request, sid, state, 1, str(exc), 400)
    set_sources(state, sources, originals=originals, demo_sources=[p.get("demo", False) for p in pending])
    return render(request, sid, state, 2)


async def parse_pending(pending, form):
    sources, originals = [], []
    for si, p in enumerate(pending):
        if "source" in p:
            sources.append(p["source"])
            originals.append(p["original"])
        else:
            selected = form.get("sheet_" + str(si))
            if len(p["sheets"]) > 1 and selected not in p["sheets"]:
                raise DataError("Виберіть аркуш для кожної книги.")
            source = await run_in_threadpool(read_source, p["name"], p["data"], p["encoding"], p["delimiter"], selected)
            sources.append(source)
            originals.append(source)
    return sources, originals


@app.post("/sheets")
async def sheets(request: Request):
    sid, state = session(request)
    form = await request.form()
    check_csrf(state, form)
    if not state.get("pending"):
        raise DataError("Завантажте книги ще раз.")
    try:
        sources, originals = await parse_pending(state["pending"], form)
    except DataError as exc:
        return render(request, sid, state, 1, str(exc), 400)
    flags = [p.get("demo", False) for p in state["pending"]]
    set_sources(state, sources, originals=originals, demo_sources=flags)
    return render(request, sid, state, 2)


@app.post("/demo")
async def demo(request: Request):
    sid, state = session(request)
    form = await request.form()
    check_csrf(state, form)
    sources = [read_source(name, (ROOT / "TEST_DATA" / name).read_bytes()) for name in ["supplier_a.csv", "supplier_b.xlsx"]]
    set_sources(state, sources, demo=True)
    return render(request, sid, state, 2)


@app.post("/analyze")
async def inspect(request: Request):
    sid, state = session(request)
    form = await request.form(max_fields=48)
    check_csrf(state, form)
    if not state["sources"]:
        return RedirectResponse("/", status_code=303)
    revision = secrets.token_urlsafe(20)
    state.update(result=None, files=None, edit_preview=None, revision=revision)
    try:
        sources = state["sources"]
        mappings = [{field: int(str(form.get(f"{slot}_{field}", "-1"))) for field in FIELDS} for slot in "ab"]
        rules = {key: form.get(key) == "on" for key in RULES}
        rules.update({key: form.get(key) == "on" if form.get("configured") else True for key in VALIDATION_RULES})
        rules["threshold"] = str(form.get("threshold", "20"))
        compare = form.get("compare") == "on"
        result = await run_in_threadpool(analyze, sources, mappings, rules, compare)
        if state["sources"] is not sources or state["revision"] != revision:
            raise DataError("Джерела змінилися в іншій вкладці. Повторіть перевірку.")
    except (DataError, ValueError, InvalidOperation) as exc:
        message = str(exc) if isinstance(exc, DataError) else "Некоректне зіставлення колонок."
        return render(request, sid, state, 2, message, 400)
    state.update(mappings=mappings, rules=rules, compare=compare, files=None, edit_preview=None,
                 revision=secrets.token_urlsafe(20))
    state["result"] = enrich_result(state, result)
    return render(request, sid, state, 3)


def check_revision(state, form):
    if not state.get("result") or not secrets.compare_digest(str(form.get("revision", "")), state.get("revision", "")):
        raise DataError("Попередній перегляд застарів. Оновіть сторінку.")


@app.post("/edit/preview")
async def edit_preview(request: Request):
    sid, state = session(request)
    form = await request.form()
    check_csrf(state, form)
    check_revision(state, form)
    try:
        slot, column = str(form.get("target", "")).split(":")
        patches = preview_edits(state["sources"], int(slot), int(column), str(form.get("operation", "")),
                                int(str(form.get("row", "2"))), str(form.get("value", "")))
        state["edit_preview"] = dict(patches=patches, token=secrets.token_urlsafe(20))
    except (ValueError, DataError) as exc:
        return render(request, sid, state, 3, str(exc) if isinstance(exc, DataError) else "Перевірте номер рядка та колонку.", 400)
    return render(request, sid, state, 3)


@app.post("/edit/{action}")
async def edit_apply(request: Request, action: str):
    sid, state = session(request)
    form = await request.form()
    check_csrf(state, form)
    check_revision(state, form)
    history = list(state.get("history", []))
    if action == "apply":
        preview = state.get("edit_preview")
        if not preview or not secrets.compare_digest(str(form.get("preview_token", "")), preview["token"]):
            raise DataError("Перегляд правок застарів. Створіть його знову.")
        patches = preview["patches"]
        sources = apply_edits(state["sources"], patches)
        if patches:
            history = (history + [patches])[-20:]
    elif action == "undo" and history:
        sources = apply_edits(state["sources"], history.pop(), reverse=True)
    elif action == "restore":
        sources, history = deepcopy(state["originals"]), []
    else:
        raise DataError("Немає доступної операції скасування.")
    revision = state["revision"]
    result = await run_in_threadpool(analyze, sources, state["mappings"], state["rules"], state["compare"])
    if revision != state["revision"]:
        raise DataError("Дані змінилися в іншій вкладці. Повторіть дію.")
    state.update(sources=sources, history=history, edit_preview=None, files=None, revision=secrets.token_urlsafe(20))
    state["result"] = enrich_result(state, result)
    return render(request, sid, state, 3)


@app.post("/apply")
async def apply(request: Request):
    sid, state = session(request)
    form = await request.form()
    check_csrf(state, form)
    if not state["result"] or not secrets.compare_digest(str(form.get("revision", "")), state.get("revision", "")):
        raise DataError("Попередній перегляд застарів. Повторіть перевірку перед експортом.")
    if form.get("acknowledge") != "on":
        return render(request, sid, state, 3, "Підтвердьте, що переглянули результат та статус перевірки.", 400)
    try:
        revision = state["revision"]
        files = await run_in_threadpool(export_files, state["result"], str(form.get("delimiter", ";")))
        if state.get("revision") != revision or not state["result"]:
            raise DataError("Результат змінився в іншій вкладці. Повторіть підтвердження.")
    except DataError as exc:
        return render(request, sid, state, 3, str(exc), 400)
    state["files"] = files
    return render(request, sid, state, 4)


@app.get("/download/{name}")
async def download(request: Request, name: str):
    _, state = session(request)
    if not state["files"] or name not in state["files"]:
        return Response("Спочатку перевірте та підтвердьте експорт.", status_code=404)
    content_type = "application/zip" if name.endswith(".zip") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if name.endswith(".xlsx") else "application/json" if name.endswith(".json") else "text/plain; charset=utf-8" if name.endswith(".txt") else "text/csv; charset=utf-8"
    return Response(state["files"][name], media_type=content_type, headers={"Content-Disposition": f'attachment; filename="{name}"'})


@app.post("/reset")
async def reset(request: Request):
    sid, state = session(request)
    check_csrf(state, await request.form())
    SESSIONS.pop(sid, None)
    response = RedirectResponse("/", status_code=303)
    response.delete_cookie("df_session")
    return response


@app.get("/health")
async def health():
    return {"status": "ok", "mode": "local"}
