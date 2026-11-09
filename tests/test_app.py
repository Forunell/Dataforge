import re
import time
from pathlib import Path

from fastapi.testclient import TestClient

from dataforge.app import SESSIONS, app

DATA = Path(__file__).resolve().parents[1] / "TEST_DATA"


def token(response, name="csrf"):
    return re.search(rf'name="{name}" value="([^"]+)"', response.text)[1]


def demo_preview(client):
    home = client.get("/")
    csrf = token(home)
    loaded = client.post("/demo", data={"csrf": csrf})
    assert loaded.status_code == 200
    mapping = {f"{s}_{k}": str(i) for s in "ab" for i, k in enumerate(["sku", "name", "price", "currency", "stock"])}
    preview = client.post("/analyze", data={"csrf": csrf, "compare": "on", **mapping})
    return csrf, mapping, preview


def test_complete_flow_and_invalidation(client):
    csrf, mapping, preview = demo_preview(client)
    assert preview.status_code == 200
    assert "Ціни не порівнюються" in preview.text
    assert client.get("/download/normalized.csv").status_code == 404
    rev = token(preview, "revision")
    assert client.post("/apply", data={"csrf": csrf, "revision": rev}).status_code == 400
    exported = client.post("/apply", data={"csrf": csrf, "revision": rev, "acknowledge": "on"})
    assert exported.status_code == 200
    assert client.get("/download/dataforge-results.zip").content.startswith(b"PK")
    assert client.get("/download/normalized.csv").status_code == 404
    assert client.get("/download/review_rows.csv").content.startswith(b"\xef\xbb\xbf")
    assert client.get("/download/unknown").status_code == 404
    for step in [1, 2, 3, 4]:
        assert client.get(f"/step/{step}").status_code == 200
    client.post("/analyze", data={"csrf": csrf, **mapping})
    assert client.get("/download/normalized.csv").status_code == 404
    assert client.post("/apply", data={"csrf": csrf, "revision": rev, "acknowledge": "on"}).status_code == 400
    reset = client.post("/reset", data={"csrf": csrf})
    assert "Чисті дані починаються тут" in reset.text


def test_upload_manual_encoding_and_failed_upload_retains_state(client):
    csrf = token(client.get("/"))
    files = {"file_a": ("cp1251.csv", (DATA / "cp1251.csv").read_bytes()), "file_b": ("b.xlsx", (DATA / "supplier_b.xlsx").read_bytes())}
    failed = client.post("/upload", data={"csrf": csrf}, files=files)
    assert failed.status_code == 400
    loaded = client.post("/upload", data={"csrf": csrf, "encoding_a": "cp1251"}, files=files)
    assert loaded.status_code == 200
    assert "000001" in loaded.text
    broken = client.post("/upload", data={"csrf": csrf})
    assert broken.status_code == 400
    assert "cp1251.csv" in client.get("/step/2").text


def test_csrf_origin_host_and_session_isolation(client):
    csrf, _, preview = demo_preview(client)
    assert client.post("/demo", data={"csrf": "wrong"}).status_code == 400
    assert client.post("/demo", data={"csrf": csrf}, headers={"Origin": "https://evil.invalid"}).status_code == 403
    assert client.post("/demo", data={"csrf": csrf}, headers={"Origin": "http://testserver"}).status_code == 200
    assert client.get("/", headers={"Host": "evil.invalid"}).status_code == 400
    assert "frame-ancestors 'none'" in preview.headers["content-security-policy"]
    with TestClient(app) as other:
        assert "Чисті дані починаються тут" in other.get("/").text
        assert other.get("/download/normalized.csv").status_code == 404


def test_html_escaping_and_invalid_mapping(client):
    csrf = token(client.get("/"))
    raw = b'sku;name;price;currency;stock\n001;<script>alert(1)</script>;1;UAH;1'
    response = client.post("/upload", data={"csrf": csrf}, files={"file_a": ("a.csv", raw), "file_b": ("b.csv", raw)})
    assert "&lt;script&gt;" in response.text
    assert "<script>alert(1)</script>" not in response.text
    assert client.post("/analyze", data={"csrf": csrf, "a_sku": "bad"}).status_code == 400


def test_request_limit_and_expiry(client):
    csrf = token(client.get("/"))
    assert client.post("/upload", content=b"x" * (4 * 1024 * 1024 + 65537)).status_code == 413
    list(SESSIONS.values())[0]["touched"] = time.monotonic() - 3601
    assert client.post("/demo", data={"csrf": csrf}).status_code == 400
