from __future__ import annotations

import hashlib
import json
import mimetypes
import re
import time
from dataclasses import dataclass, field
from typing import Any, Callable

import requests

from crm_store import (
    CrmConfigurationError,
    build_normalized_lead_payload,
    create_store,
    normalize_amo_domain,
    normalize_amo_token,
)


class AmoCrmDeliveryError(RuntimeError):
    pass


@dataclass(frozen=True)
class DeliveryArtifact:
    filename: str
    data: bytes
    content_type: str = "application/octet-stream"


@dataclass
class AmoDeliveryResult:
    status: str
    message: str
    lead_id: int | None = None
    contact_id: int | None = None
    company_id: int | None = None
    task_id: int | None = None
    attached_files: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _int_setting(settings: dict[str, Any], key: str) -> int:
    raw = str(settings.get(key, "") or "").strip()
    if not raw.isdigit() or int(raw) <= 0:
        raise CrmConfigurationError(f"В настройках amoCRM не указан корректный {key}.")
    return int(raw)


def _normalize_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def _field_values(entity: dict[str, Any], field_code: str) -> list[str]:
    values: list[str] = []
    for item in entity.get("custom_fields_values") or []:
        if str(item.get("field_code", "")).upper() != field_code.upper():
            continue
        for value in item.get("values") or []:
            raw = value.get("value") if isinstance(value, dict) else value
            if raw not in (None, ""):
                values.append(str(raw))
    return values


def build_delivery_idempotency_key(
    payload: dict[str, Any],
    artifacts: list[DeliveryArtifact],
) -> str:
    identity = {
        "company": _normalize_text(payload.get("company")),
        "email": _normalize_text(payload.get("email")),
        "phone": re.sub(r"\D+", "", str(payload.get("phone", ""))),
        "source_app": str(payload.get("source_app", "")),
        "audit_day": str(payload.get("created_at", ""))[:10],
        "security_maturity": int(payload.get("security_maturity", 0) or 0),
        "it_maturity": int(payload.get("it_maturity", 0) or 0),
    }
    digest = hashlib.sha256(
        json.dumps(identity, sort_keys=True, ensure_ascii=False).encode("utf-8")
    )
    return f"amocrm:{digest.hexdigest()}"


class AmoCrmClient:
    def __init__(
        self,
        domain: str,
        access_token: str,
        *,
        timeout: int = 25,
        session: requests.Session | None = None,
    ):
        self.domain = normalize_amo_domain(domain)
        self.access_token = normalize_amo_token(access_token)
        if not self.access_token:
            raise CrmConfigurationError("Не задан access token amoCRM.")
        self.timeout = timeout
        self.session = session or requests.Session()

    @property
    def api_base(self) -> str:
        return f"https://{self.domain}"

    @property
    def headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def _request(
        self,
        method: str,
        path_or_url: str,
        *,
        params: dict[str, Any] | None = None,
        payload: Any = None,
        data: bytes | None = None,
        headers: dict[str, str] | None = None,
        expected: tuple[int, ...] = (200,),
    ) -> Any:
        url = (
            path_or_url
            if path_or_url.startswith("https://")
            else f"{self.api_base}{path_or_url}"
        )
        request_headers = dict(self.headers)
        if headers:
            request_headers.update(headers)
        try:
            response = self.session.request(
                method,
                url,
                params=params,
                json=payload if data is None else None,
                data=data,
                headers=request_headers,
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise AmoCrmDeliveryError(f"amoCRM недоступна: {exc}") from exc
        if response.status_code not in expected:
            detail = (response.text or "")[:600].replace(self.access_token, "***")
            raise AmoCrmDeliveryError(
                f"amoCRM HTTP {response.status_code} для {method} {url}: {detail}"
            )
        if not response.content:
            return {}
        try:
            return response.json()
        except ValueError as exc:
            raise AmoCrmDeliveryError(
                f"amoCRM вернула некорректный JSON для {method} {url}."
            ) from exc

    @staticmethod
    def _first_embedded(response: dict[str, Any], key: str) -> dict[str, Any]:
        values = (response.get("_embedded") or {}).get(key) or []
        if not values:
            raise AmoCrmDeliveryError(f"amoCRM не вернула созданную сущность {key}.")
        return values[0]

    def find_company(self, name: str) -> dict[str, Any] | None:
        response = self._request(
            "GET",
            "/api/v4/companies",
            params={"query": name, "limit": 50},
        )
        companies = (response.get("_embedded") or {}).get("companies") or []
        expected_name = _normalize_text(name)
        return next(
            (company for company in companies if _normalize_text(company.get("name")) == expected_name),
            None,
        )

    def ensure_company(self, name: str, responsible_user_id: int) -> tuple[int, bool]:
        existing = self.find_company(name)
        if existing:
            return int(existing["id"]), False
        response = self._request(
            "POST",
            "/api/v4/companies",
            payload=[{"name": name, "responsible_user_id": responsible_user_id}],
        )
        return int(self._first_embedded(response, "companies")["id"]), True

    def find_contact(self, email: str, phone: str) -> dict[str, Any] | None:
        searches = [value for value in (email, phone) if value]
        normalized_email = _normalize_text(email)
        normalized_phone = re.sub(r"\D+", "", phone)
        for query in searches:
            response = self._request(
                "GET",
                "/api/v4/contacts",
                params={"query": query, "limit": 50, "with": "companies"},
            )
            contacts = (response.get("_embedded") or {}).get("contacts") or []
            for contact in contacts:
                emails = {_normalize_text(value) for value in _field_values(contact, "EMAIL")}
                phones = {
                    re.sub(r"\D+", "", value)
                    for value in _field_values(contact, "PHONE")
                }
                if normalized_email and normalized_email in emails:
                    return contact
                if normalized_phone and normalized_phone in phones:
                    return contact
        return None

    def ensure_contact(
        self,
        *,
        name: str,
        email: str,
        phone: str,
        company_id: int,
        responsible_user_id: int,
    ) -> tuple[int, bool]:
        existing = self.find_contact(email, phone)
        if existing:
            contact_id = int(existing["id"])
            companies = (existing.get("_embedded") or {}).get("companies") or []
            if not any(int(item.get("id", 0) or 0) == company_id for item in companies):
                self._request(
                    "POST",
                    f"/api/v4/contacts/{contact_id}/link",
                    payload=[
                        {
                            "to_entity_id": company_id,
                            "to_entity_type": "companies",
                        }
                    ],
                    expected=(200, 204),
                )
            return contact_id, False

        custom_fields = []
        if email:
            custom_fields.append(
                {
                    "field_code": "EMAIL",
                    "values": [{"value": email, "enum_code": "WORK"}],
                }
            )
        if phone:
            custom_fields.append(
                {
                    "field_code": "PHONE",
                    "values": [{"value": phone, "enum_code": "WORK"}],
                }
            )
        model: dict[str, Any] = {
            "name": name or "Контакт автоматического аудита",
            "responsible_user_id": responsible_user_id,
            "_embedded": {"companies": [{"id": company_id}]},
        }
        if custom_fields:
            model["custom_fields_values"] = custom_fields
        response = self._request("POST", "/api/v4/contacts", payload=[model])
        return int(self._first_embedded(response, "contacts")["id"]), True

    def create_lead(
        self,
        *,
        company_name: str,
        company_id: int,
        contact_id: int,
        pipeline_id: int,
        status_id: int,
        responsible_user_id: int,
    ) -> int:
        title = f"Автоматический аудит — {company_name}"
        response = self._request(
            "POST",
            "/api/v4/leads",
            payload=[
                {
                    "name": title,
                    "pipeline_id": pipeline_id,
                    "status_id": status_id,
                    "responsible_user_id": responsible_user_id,
                    "_embedded": {
                        "contacts": [{"id": contact_id, "is_main": True}],
                        "companies": [{"id": company_id}],
                        "tags": [{"name": "Автоматический аудит"}],
                    },
                }
            ],
        )
        return int(self._first_embedded(response, "leads")["id"])

    def create_task(
        self,
        *,
        lead_id: int,
        company_name: str,
        responsible_user_id: int,
        due_hours: int,
    ) -> int:
        due_at = int(time.time()) + max(1, due_hours) * 3600
        response = self._request(
            "POST",
            "/api/v4/tasks",
            payload=[
                {
                    "task_type_id": 1,
                    "text": f"Автоматический аудит — {company_name}",
                    "complete_till": due_at,
                    "entity_id": lead_id,
                    "entity_type": "leads",
                    "responsible_user_id": responsible_user_id,
                }
            ],
        )
        return int(self._first_embedded(response, "tasks")["id"])

    def get_drive_url(self) -> str:
        response = self._request(
            "GET",
            "/api/v4/account",
            params={"with": "drive_url"},
        )
        drive_url = str(
            response.get("drive_url")
            or (response.get("_embedded") or {}).get("drive_url")
            or ""
        ).rstrip("/")
        if not drive_url.startswith("https://"):
            raise AmoCrmDeliveryError(
                "amoCRM не вернула drive_url. Проверьте доступ интеграции к файлам."
            )
        return drive_url

    def upload_file(self, artifact: DeliveryArtifact, drive_url: str) -> str:
        if not artifact.data:
            raise AmoCrmDeliveryError(f"Файл {artifact.filename} пуст.")
        content_type = (
            artifact.content_type
            or mimetypes.guess_type(artifact.filename)[0]
            or "application/octet-stream"
        )
        session = self._request(
            "POST",
            f"{drive_url}/v1.0/sessions",
            payload={
                "file_name": artifact.filename,
                "file_size": len(artifact.data),
                "content_type": content_type,
            },
        )
        upload_url = str(session.get("upload_url") or "")
        part_size = int(session.get("max_part_size") or 0)
        if not upload_url or part_size <= 0:
            raise AmoCrmDeliveryError(
                f"amoCRM не открыла сессию загрузки для {artifact.filename}."
            )
        final_response: dict[str, Any] = {}
        for offset in range(0, len(artifact.data), part_size):
            chunk = artifact.data[offset : offset + part_size]
            final_response = self._request(
                "POST",
                upload_url,
                data=chunk,
                headers={"Content-Type": "application/octet-stream"},
            )
            if offset + len(chunk) < len(artifact.data):
                upload_url = str(final_response.get("next_url") or "")
                if not upload_url:
                    raise AmoCrmDeliveryError(
                        f"amoCRM не вернула URL следующей части {artifact.filename}."
                    )
        file_uuid = str(final_response.get("uuid") or "")
        if not file_uuid:
            raise AmoCrmDeliveryError(f"amoCRM не вернула UUID файла {artifact.filename}.")
        return file_uuid

    def attach_file_to_lead(self, lead_id: int, file_uuid: str) -> None:
        self._request(
            "PUT",
            f"/api/v4/leads/{lead_id}/files",
            payload=[{"file_uuid": file_uuid}],
            expected=(202,),
        )

    def deliver(
        self,
        settings: dict[str, Any],
        payload: dict[str, Any],
        artifacts: list[DeliveryArtifact],
    ) -> AmoDeliveryResult:
        company_name = str(payload.get("company") or "").strip()
        if not company_name:
            raise CrmConfigurationError("Для CRM не указано название компании.")
        responsible_user_id = _int_setting(settings, "responsible_user_id")
        pipeline_id = _int_setting(settings, "pipeline_id")
        status_id = _int_setting(settings, "status_id")
        due_hours = max(1, min(720, int(settings.get("task_due_hours", 24) or 24)))

        company_id, _ = self.ensure_company(company_name, responsible_user_id)
        contact_id, _ = self.ensure_contact(
            name=str(payload.get("contact_name") or "").strip(),
            email=str(payload.get("email") or "").strip(),
            phone=str(payload.get("phone") or "").strip(),
            company_id=company_id,
            responsible_user_id=responsible_user_id,
        )
        lead_id = self.create_lead(
            company_name=company_name,
            company_id=company_id,
            contact_id=contact_id,
            pipeline_id=pipeline_id,
            status_id=status_id,
            responsible_user_id=responsible_user_id,
        )
        task_id = self.create_task(
            lead_id=lead_id,
            company_name=company_name,
            responsible_user_id=responsible_user_id,
            due_hours=due_hours,
        )

        result = AmoDeliveryResult(
            status="success",
            message=f"Создана сделка #{lead_id}: Автоматический аудит — {company_name}",
            lead_id=lead_id,
            contact_id=contact_id,
            company_id=company_id,
            task_id=task_id,
        )
        if artifacts:
            try:
                drive_url = self.get_drive_url()
                for artifact in artifacts:
                    try:
                        file_uuid = self.upload_file(artifact, drive_url)
                        self.attach_file_to_lead(lead_id, file_uuid)
                        result.attached_files.append(artifact.filename)
                    except (AmoCrmDeliveryError, CrmConfigurationError) as exc:
                        result.warnings.append(f"{artifact.filename}: {exc}")
            except (AmoCrmDeliveryError, CrmConfigurationError) as exc:
                result.warnings.append(str(exc))
        if result.warnings:
            result.status = "partial"
            result.message += "; вложения загружены не полностью"
        return result


def deliver_audit_to_active_crm(
    secret_getter: Callable[[str, Any], Any],
    runtime_settings: dict[str, Any],
    *,
    client_info: dict[str, Any],
    security_maturity: int,
    it_maturity: int,
    source_app: str,
    priorities: list[dict[str, Any]] | None,
    artifacts: list[DeliveryArtifact],
) -> AmoDeliveryResult:
    provider = str(runtime_settings.get("active_provider", "off") or "off").lower()
    if provider == "off":
        return AmoDeliveryResult("skipped", "CRM-интеграция выключена.")
    if provider != "amocrm":
        return AmoDeliveryResult(
            "skipped",
            f"Автоматическая отправка для {provider} пока не реализована.",
        )

    try:
        store = create_store(secret_getter)
        config = store.get_provider_config("amocrm")
        if (
            not config
            or config.get("connection_status") != "ok"
            or not bool(config.get("has_secret"))
        ):
            return AmoDeliveryResult("error", "Активная конфигурация amoCRM не найдена.")
        settings = config.get("settings") if isinstance(config.get("settings"), dict) else {}
        credentials = store.get_provider_credentials("amocrm")
        payload = build_normalized_lead_payload(
            client_info,
            security_maturity,
            it_maturity,
            source_app,
            priorities,
        )
        idempotency_key = build_delivery_idempotency_key(payload, artifacts)
        existing = store.get_delivery_by_idempotency(idempotency_key)
        if existing:
            return AmoDeliveryResult(
                "skipped",
                (
                    "Этот результат аудита уже отправлялся в amoCRM: "
                    f"{existing.get('lead_reference') or existing.get('status') or 'запись найдена'}."
                ),
            )
        if not store.reserve_delivery("amocrm", "audit_completed", idempotency_key):
            return AmoDeliveryResult("skipped", "Отправка этого аудита уже выполняется.")

        client = AmoCrmClient(
            settings.get("domain", ""),
            credentials.get("access_token", ""),
        )
        result = client.deliver(settings, payload, artifacts)
        lead_reference = (
            f"https://{client.domain}/leads/detail/{result.lead_id}"
            if result.lead_id
            else None
        )
        log_message = result.message
        if result.warnings:
            log_message = f"{log_message}. {' | '.join(result.warnings)}"
        store.update_delivery(
            idempotency_key,
            status=result.status,
            message=log_message,
            lead_reference=lead_reference,
        )
        return result
    except Exception as exc:
        message = f"Ошибка отправки в amoCRM: {exc}"
        try:
            if "store" in locals() and "idempotency_key" in locals():
                store.update_delivery(
                    idempotency_key,
                    status="error",
                    message=message,
                )
        except Exception:
            pass
        return AmoDeliveryResult("error", message)
