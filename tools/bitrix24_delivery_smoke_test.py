from __future__ import annotations

import sys
import types
from pathlib import Path


requests_stub = types.ModuleType("requests")
requests_stub.RequestException = Exception
requests_stub.Session = object
requests_stub.get = lambda *args, **kwargs: None
requests_stub.post = lambda *args, **kwargs: None
requests_stub.request = lambda *args, **kwargs: None
sys.modules.setdefault("requests", requests_stub)
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import crm_store  # noqa: E402
from crm_delivery import Bitrix24Client, DeliveryArtifact  # noqa: E402


class FakeResponse:
    def __init__(self, payload, status_code=200):
        self.status_code = status_code
        self._payload = payload
        self.content = b"json"
        self.text = str(payload)

    def json(self):
        return self._payload


class FakeSession:
    def __init__(self):
        self.calls = []

    def post(self, url, **kwargs):
        self.calls.append((url, kwargs.get("json") or {}))
        method = url.rsplit("/", 1)[-1].removesuffix(".json")
        responses = {
            "crm.company.list": [],
            "crm.company.add": 101,
            "crm.duplicate.findbycomm": {},
            "crm.contact.add": 202,
            "crm.deal.add": 303,
            "tasks.task.add": {"task": {"id": 404}},
            "crm.timeline.comment.add": 505,
        }
        if method not in responses:
            raise AssertionError(f"Unexpected Bitrix24 method: {method}")
        return FakeResponse({"result": responses[method]})


def main():
    session = FakeSession()
    client = Bitrix24Client(
        "https://example.bitrix24.kz/rest/7/test-secret/",
        session=session,
    )
    result = client.deliver(
        {
            "category_id": "0",
            "stage_id": "C0:NEW",
            "responsible_user_id": "7",
            "task_due_hours": 24,
        },
        {
            "company": "Demo Company",
            "contact_name": "Ivan Petrov",
            "contact_role": "ИТ-директор",
            "email": "ivan@example.kz",
            "phone": "+7 777 000 00 00",
            "source_app": "Test",
            "security_maturity": 61,
            "it_maturity": 54,
        },
        [
            DeliveryArtifact("audit.pptx", b"presentation"),
            DeliveryArtifact("sales.xlsx", b"playbook"),
        ],
    )

    assert result.status == "success"
    assert (result.company_id, result.contact_id, result.lead_id, result.task_id) == (
        101,
        202,
        303,
        404,
    )
    assert result.attached_files == ["audit.pptx", "sales.xlsx"]

    calls = {method.rsplit("/", 1)[-1].removesuffix(".json"): payload for method, payload in session.calls}
    company = calls["crm.company.add"]["fields"]
    assert company["TITLE"] == "Demo Company"
    assert company["PHONE"][0]["VALUE"] == "+7 777 000 00 00"
    assert company["EMAIL"][0]["VALUE"] == "ivan@example.kz"

    contact = calls["crm.contact.add"]["fields"]
    assert contact["COMPANY_ID"] == 101
    assert contact["POST"] == "ИТ-директор"

    deal = calls["crm.deal.add"]["fields"]
    assert deal["TITLE"] == "[Test] Автоматический аудит — Demo Company"
    assert deal["CATEGORY_ID"] == 0
    assert deal["STAGE_ID"] == "C0:NEW"

    task = calls["tasks.task.add"]["fields"]
    assert task["RESPONSIBLE_ID"] == 7
    assert task["UF_CRM_TASK"] == ["D_303"]

    timeline_calls = [
        payload
        for method, payload in session.calls
        if method.endswith("/crm.timeline.comment.add.json")
    ]
    assert len(timeline_calls) == 2
    assert timeline_calls[0]["fields"]["FILES"][0][0] == "audit.pptx"
    assert timeline_calls[1]["fields"]["FILES"][0][0] == "sales.xlsx"

    def fake_connection_post(url, **kwargs):
        method = url.rsplit("/", 1)[-1].removesuffix(".json")
        responses = {
            "profile": {"ID": "7", "NAME": "Ivan", "LAST_NAME": "Rudoy"},
            "user.get": [{"ID": "7", "NAME": "Ivan", "LAST_NAME": "Rudoy"}],
            "crm.category.list": {
                "categories": [
                    {
                        "id": 0,
                        "name": "Основная",
                        "entityTypeId": 2,
                        "isDefault": "Y",
                    }
                ]
            },
            "crm.status.list": [
                {
                    "ENTITY_ID": "DEAL_STAGE",
                    "STATUS_ID": "NEW",
                    "NAME": "Новая",
                }
            ],
            "tasks.task.list": {"tasks": []},
        }
        if method not in responses:
            raise AssertionError(f"Unexpected connection method: {method}")
        return FakeResponse({"result": responses[method]})

    crm_store.requests.post = fake_connection_post
    connection = crm_store.test_bitrix_connection(
        {
            "category_id": "0",
            "stage_id": "NEW",
            "responsible_user_id": "7",
        },
        {"webhook_url": "https://example.bitrix24.kz/rest/7/test-secret/"},
    )
    assert connection.ok is True
    assert connection.details["responsible_name"] == "Ivan Rudoy"

    print("Bitrix24 delivery smoke test: OK")


if __name__ == "__main__":
    main()
