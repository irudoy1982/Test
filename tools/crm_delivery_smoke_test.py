from __future__ import annotations

import sys
import types
from pathlib import Path


requests_stub = types.ModuleType("requests")
requests_stub.RequestException = Exception
requests_stub.Session = object
requests_stub.get = lambda *args, **kwargs: None
requests_stub.request = lambda *args, **kwargs: None
sys.modules.setdefault("requests", requests_stub)
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from crm_delivery import AmoCrmClient, DeliveryArtifact  # noqa: E402


class FakeResponse:
    def __init__(self, status_code, payload=None):
        self.status_code = status_code
        self._payload = payload
        self.content = b"" if payload is None else b"json"
        self.text = "" if payload is None else str(payload)

    def json(self):
        return self._payload


class FakeSession:
    def __init__(self):
        self.calls = []

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        if method == "GET" and url.endswith("/api/v4/companies"):
            return FakeResponse(200, {"_embedded": {"companies": []}})
        if method == "POST" and url.endswith("/api/v4/companies"):
            return FakeResponse(200, {"_embedded": {"companies": [{"id": 101}]}})
        if method == "GET" and url.endswith("/api/v4/contacts"):
            return FakeResponse(200, {"_embedded": {"contacts": []}})
        if method == "POST" and url.endswith("/api/v4/contacts"):
            return FakeResponse(200, {"_embedded": {"contacts": [{"id": 202}]}})
        if method == "POST" and url.endswith("/api/v4/leads"):
            return FakeResponse(200, {"_embedded": {"leads": [{"id": 303}]}})
        if method == "POST" and url.endswith("/api/v4/tasks"):
            return FakeResponse(200, {"_embedded": {"tasks": [{"id": 404}]}})
        if method == "GET" and url.endswith("/api/v4/account"):
            return FakeResponse(200, {"drive_url": "https://drive.example"})
        if method == "POST" and url.endswith("/v1.0/sessions"):
            return FakeResponse(
                200,
                {
                    "upload_url": "https://drive.example/upload/part-1",
                    "max_part_size": 1024,
                },
            )
        if method == "POST" and "/upload/" in url:
            return FakeResponse(200, {"uuid": "file-uuid"})
        if method == "PUT" and url.endswith("/api/v4/leads/303/files"):
            return FakeResponse(202)
        raise AssertionError(f"Unexpected request: {method} {url}")


def main():
    session = FakeSession()
    client = AmoCrmClient(
        "example.amocrm.ru",
        "test-access-token",
        session=session,
    )
    result = client.deliver(
        {
            "pipeline_id": "11",
            "status_id": "22",
            "responsible_user_id": "33",
            "task_due_hours": 24,
        },
        {
            "company": "Demo Company",
            "contact_name": "Ivan Petrov",
            "email": "ivan@example.kz",
            "phone": "+7 777 000 00 00",
        },
        [DeliveryArtifact("audit.pptx", b"presentation")],
    )

    assert result.status == "success"
    assert (result.company_id, result.contact_id, result.lead_id, result.task_id) == (
        101,
        202,
        303,
        404,
    )
    assert result.attached_files == ["audit.pptx"]

    lead_call = next(call for call in session.calls if call[1].endswith("/api/v4/leads"))
    lead = lead_call[2]["json"][0]
    assert lead["name"] == "Автоматический аудит — Demo Company"
    assert lead["pipeline_id"] == 11
    assert lead["status_id"] == 22
    assert lead["_embedded"]["contacts"][0]["is_main"] is True

    task_call = next(call for call in session.calls if call[1].endswith("/api/v4/tasks"))
    task = task_call[2]["json"][0]
    assert task["text"] == "Автоматический аудит — Demo Company"
    assert task["responsible_user_id"] == 33

    print("CRM delivery smoke test: OK")


if __name__ == "__main__":
    main()
