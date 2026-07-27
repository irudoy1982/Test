from __future__ import annotations

import sys
import types
from datetime import datetime, timedelta, timezone
from pathlib import Path


requests_stub = types.ModuleType("requests")
requests_stub.RequestException = Exception
requests_stub.Session = object
requests_stub.get = lambda *args, **kwargs: None
requests_stub.request = lambda *args, **kwargs: None
sys.modules.setdefault("requests", requests_stub)
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import crm_delivery  # noqa: E402
from crm_delivery import AmoCrmClient, AmoDeliveryResult, DeliveryArtifact  # noqa: E402


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
        self.upload_part = 0

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        if method == "GET" and url.endswith("/api/v4/companies"):
            return FakeResponse(204)
        if method == "POST" and url.endswith("/api/v4/companies"):
            return FakeResponse(200, {"_embedded": {"companies": [{"id": 101}]}})
        if method == "GET" and url.endswith("/api/v4/contacts"):
            return FakeResponse(204)
        if method == "POST" and url.endswith("/api/v4/contacts"):
            return FakeResponse(200, {"_embedded": {"contacts": [{"id": 202}]}})
        if method == "POST" and url.endswith("/api/v4/contacts/202/link"):
            return FakeResponse(204)
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
                    "max_part_size": 8,
                },
            )
        if method == "POST" and "/upload/" in url:
            self.upload_part += 1
            if self.upload_part == 1:
                return FakeResponse(
                    202,
                    {"next_url": "https://drive.example/upload/part-2"},
                )
            return FakeResponse(201, {"uuid": "file-uuid"})
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
            "source_app": "Test",
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

    company_call = next(
        call for call in session.calls if call[1].endswith("/api/v4/companies")
        and call[0] == "POST"
    )
    company_fields = company_call[2]["json"][0]["custom_fields_values"]
    assert {item["field_code"] for item in company_fields} == {"EMAIL", "PHONE"}

    lead_call = next(call for call in session.calls if call[1].endswith("/api/v4/leads"))
    lead = lead_call[2]["json"][0]
    assert lead["name"] == "[Test] Автоматический аудит — Demo Company"
    assert lead["pipeline_id"] == 11
    assert lead["status_id"] == 22
    assert lead["_embedded"]["contacts"][0]["is_main"] is True

    task_call = next(call for call in session.calls if call[1].endswith("/api/v4/tasks"))
    task = task_call[2]["json"][0]
    assert task["text"] == "[Test] Автоматический аудит — Demo Company"
    assert task["responsible_user_id"] == 33

    class FakeStore:
        def __init__(self, existing_status="error", created_at=None):
            self.status_updates = []
            self.existing_status = existing_status
            self.created_at = created_at or (
                datetime.now(timezone.utc) - timedelta(minutes=10)
            ).isoformat()

        def get_provider_config(self, provider):
            return {
                "provider": provider,
                "settings": {
                    "domain": "example.amocrm.ru",
                    "pipeline_id": "11",
                    "status_id": "22",
                    "responsible_user_id": "33",
                },
                "has_secret": True,
                "connection_status": "ok",
            }

        def get_provider_credentials(self, provider):
            return {"access_token": "test-access-token"}

        def get_delivery_by_idempotency(self, key):
            return {
                "status": self.existing_status,
                "created_at": self.created_at,
                "lead_reference": "https://example.amocrm.ru/leads/detail/303",
            }

        def reserve_delivery(self, provider, event, key):
            return True

        def update_delivery(self, key, **values):
            self.status_updates.append(values.get("status"))
            return None

    class FakeClient:
        domain = "example.amocrm.ru"
        repaired_lead_id = None
        existing_lead = True

        def __init__(self, domain, token):
            pass

        def deliver(self, settings, payload, artifacts):
            return AmoDeliveryResult(
                status="success",
                message="ok",
                lead_id=303,
            )

        def repair_lead_relationships(self, lead_id):
            self.repaired_lead_id = lead_id

        def update_lead_title(self, lead_id, company_name, source_app):
            self.updated_title = (
                lead_id,
                f"[{source_app}] Автоматический аудит — {company_name}",
            )

        def lead_exists(self, lead_id):
            return self.existing_lead

        def attach_artifacts_to_lead(self, lead_id, artifacts):
            return AmoDeliveryResult(
                status="success",
                message="attachments restored",
                lead_id=lead_id,
                attached_files=[item.filename for item in artifacts],
            )

    original_create_store = crm_delivery.create_store
    original_client = crm_delivery.AmoCrmClient
    fake_store = FakeStore()
    crm_delivery.create_store = lambda secret_getter: fake_store
    crm_delivery.AmoCrmClient = FakeClient
    try:
        orchestrated = crm_delivery.deliver_audit_to_active_crm(
            lambda name, default=None: default,
            {"active_provider": "amocrm"},
            client_info={"Наименование компании": "Demo Company"},
            security_maturity=50,
            it_maturity=60,
            source_app="Test",
            priorities=[],
            artifacts=[],
        )
        assert orchestrated.status == "success"
        assert orchestrated.lead_id == 303
        assert fake_store.status_updates == ["pending", "success"]

        partial_store = FakeStore(existing_status="partial")
        crm_delivery.create_store = lambda secret_getter: partial_store
        recovered = crm_delivery.deliver_audit_to_active_crm(
            lambda name, default=None: default,
            {"active_provider": "amocrm"},
            client_info={"Наименование компании": "Demo Company"},
            security_maturity=50,
            it_maturity=60,
            source_app="Test",
            priorities=[],
            artifacts=[DeliveryArtifact("audit.pptx", b"presentation")],
        )
        assert recovered.status == "success"
        assert recovered.lead_id == 303
        assert recovered.attached_files == ["audit.pptx"]
        assert partial_store.status_updates == ["success"]

        deleted_store = FakeStore(existing_status="partial")
        crm_delivery.create_store = lambda secret_getter: deleted_store
        FakeClient.existing_lead = False
        recreated = crm_delivery.deliver_audit_to_active_crm(
            lambda name, default=None: default,
            {"active_provider": "amocrm"},
            client_info={"Наименование компании": "Demo Company"},
            security_maturity=50,
            it_maturity=60,
            source_app="Test",
            priorities=[],
            artifacts=[],
        )
        assert recreated.status == "success"
        assert deleted_store.status_updates == ["pending", "success"]

        deleted_success_store = FakeStore(existing_status="success")
        crm_delivery.create_store = lambda secret_getter: deleted_success_store
        FakeClient.existing_lead = False
        recreated_from_success = crm_delivery.deliver_audit_to_active_crm(
            lambda name, default=None: default,
            {"active_provider": "amocrm"},
            client_info={"Наименование компании": "Demo Company"},
            security_maturity=50,
            it_maturity=60,
            source_app="Test",
            priorities=[],
            artifacts=[],
        )
        assert recreated_from_success.status == "success"
        assert deleted_success_store.status_updates == ["pending", "success"]

        stale_pending_store = FakeStore(existing_status="pending")
        crm_delivery.create_store = lambda secret_getter: stale_pending_store
        recreated_from_pending = crm_delivery.deliver_audit_to_active_crm(
            lambda name, default=None: default,
            {"active_provider": "amocrm"},
            client_info={"Наименование компании": "Demo Company"},
            security_maturity=50,
            it_maturity=60,
            source_app="Test",
            priorities=[],
            artifacts=[],
        )
        assert recreated_from_pending.status == "success"
        assert stale_pending_store.status_updates == ["pending", "success"]
    finally:
        crm_delivery.create_store = original_create_store
        crm_delivery.AmoCrmClient = original_client

    print("CRM delivery smoke test: OK")


if __name__ == "__main__":
    main()
