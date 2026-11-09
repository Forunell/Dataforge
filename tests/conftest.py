import pytest
from fastapi.testclient import TestClient

from dataforge.app import SESSIONS, app


@pytest.fixture
def client():
    SESSIONS.clear()
    with TestClient(app) as instance:
        yield instance
    SESSIONS.clear()
