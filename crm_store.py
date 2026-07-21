from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Callable
from urllib.parse import urlparse

import requests


DEFAULT_RUNTIME_SETTINGS = {
    "active_provider": "off",
    "customer_delivery_format": "pptx",
    "test_mode": True,
    "telegram_diagnostics_enabled": True,
    "telegram_send_lead_summary": True,
    "telegram_send_sales_playbook": True,
    "telegram_send_customer_report": True,
    "telegram_lead_template": "",
    "telegram_sales_caption": "[{app}] Sales playbook: {company}",
    "telegram_customer_caption": "[{app}] Клиентское заключение: {company}",
}
ALLOWED_PROVIDERS = {"off", "amocrm", "bitrix24"}
ALLOWED_DELIVERY_FORMATS = {"pptx", "xlsx", "both"}


class CrmConfigurationError(RuntimeError):
    pass


@dataclass(frozen=True)
class ConnectionCheck:
    ok: bool
    message: str
    details: dict[str, Any]


def normalize_runtime_settings(value: Any) -> dict[str, Any]:
    source = value if isinstance(value, dict) else {}
    provider = str(source.get("active_provider", "off") or "off").lower()
    delivery_format = str(source.get("customer_delivery_format", "pptx") or "pptx").lower()
    return {
        "active_provider": provider if provider in ALLOWED_PROVIDERS else "off",
        "customer_delivery_format": (
            delivery_format if delivery_format in ALLOWED_DELIVERY_FORMATS else "pptx"
        ),
        "test_mode": bool(source.get("test_mode", True)),
        "telegram_diagnostics_enabled": bool(source.get("telegram_diagnostics_enabled", True)),
        "telegram_send_lead_summary": bool(source.get("telegram_send_lead_summary", True)),
        "telegram_send_sales_playbook": bool(source.get("telegram_send_sales_playbook", True)),
        "telegram_send_customer_report": bool(source.get("telegram_send_customer_report", True)),
        "telegram_lead_template": str(source.get("telegram_lead_template", "") or "")[:3500],
        "telegram_sales_caption": str(
            source.get("telegram_sales_caption", DEFAULT_RUNTIME_SETTINGS["telegram_sales_caption"])
            or DEFAULT_RUNTIME_SETTINGS["telegram_sales_caption"]
        )[:900],
        "telegram_customer_caption": str(
            source.get("telegram_customer_caption", DEFAULT_RUNTIME_SETTINGS["telegram_customer_caption"])
            or DEFAULT_RUNTIME_SETTINGS["telegram_customer_caption"]
        )[:900],
    }


def normalize_phone(value: Any) -> str:
    digits = re.sub(r"\D+", "", str(value or ""))
    return f"+{digits}" if digits else ""


def build_normalized_lead_payload(
    client_info: dict[str, Any],
    security_maturity: int,
    it_maturity: int,
    source_app: str,
    priorities: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    return {
        "schema": "audit-crm-lead-v1",
        "source_app": str(source_app or "Test"),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "company": str(client_info.get("Наименование компании", "") or "").strip(),
        "industry": str(client_info.get("Сфера деятельности", "") or "").strip(),
        "city": str(client_info.get("Город", "") or "").strip(),
        "website": str(client_info.get("Сайт компании", "") or "").strip(),
        "contact_name": str(client_info.get("ФИО контактного лица", "") or "").strip(),
        "contact_role": str(client_info.get("Должность", "") or "").strip(),
        "email": str(client_info.get("Email", "") or "").strip().lower(),
        "phone": normalize_phone(client_info.get("Контактный телефон", "")),
        "security_maturity": max(0, min(100, int(security_maturity or 0))),
        "it_maturity": max(0, min(100, int(it_maturity or 0))),
        "priorities": list(priorities or []),
    }


class SupabaseCrmStore:
    def __init__(self, project_url: str, secret_key: str, timeout: int = 15):
        self.project_url = str(project_url or "").strip().rstrip("/")
        self.secret_key = str(secret_key or "").strip()
        self.timeout = timeout
        if not self.project_url or not self.secret_key:
            raise CrmConfigurationError("Хранилище CRM не настроено.")
        parsed = urlparse(self.project_url)
        if parsed.scheme != "https" or not parsed.netloc:
            raise CrmConfigurationError("SUPABASE_URL должен быть корректным HTTPS-адресом.")

    @property
    def headers(self) -> dict[str, str]:
        return {
            "apikey": self.secret_key,
            "Authorization": f"Bearer {self.secret_key}",
            "Content-Type": "application/json",
        }

    def _request(
        self,
        method: str,
        path: str,
        *,
        payload: Any = None,
        params: dict[str, Any] | None = None,
        prefer: str | None = None,
    ) -> Any:
        headers = dict(self.headers)
        if prefer:
            headers["Prefer"] = prefer
        try:
            response = requests.request(
                method,
                f"{self.project_url}{path}",
                headers=headers,
                json=payload,
                params=params,
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise CrmConfigurationError(f"Хранилище CRM недоступно: {exc}") from exc
        if response.status_code >= 400:
            message = response.text[:300].replace(self.secret_key, "***")
            raise CrmConfigurationError(
                f"Ошибка хранилища CRM HTTP {response.status_code}: {message}"
            )
        if not response.content:
            return None
        try:
            return response.json()
        except ValueError:
            return response.text

    def get_runtime_settings(self) -> dict[str, Any]:
        rows = self._request(
            "GET",
            "/rest/v1/app_settings",
            params={"key": "eq.runtime", "select": "value", "limit": "1"},
        )
        if not rows:
            return dict(DEFAULT_RUNTIME_SETTINGS)
        return normalize_runtime_settings(rows[0].get("value"))

    def save_runtime_settings(self, settings: dict[str, Any], updated_by: str) -> dict[str, Any]:
        normalized = normalize_runtime_settings(settings)
        self._request(
            "POST",
            "/rest/v1/app_settings",
            payload={
                "key": "runtime",
                "value": normalized,
                "updated_by": str(updated_by or "admin"),
            },
            prefer="resolution=merge-duplicates,return=minimal",
        )
        return normalized

    def get_provider_config(self, provider: str) -> dict[str, Any]:
        provider = str(provider or "").lower()
        rows = self._request(
            "GET",
            "/rest/v1/crm_provider_configs",
            params={
                "provider": f"eq.{provider}",
                "select": (
                    "provider,settings,has_secret,connection_status,"
                    "connection_message,connection_checked_at,updated_at,updated_by"
                ),
                "limit": "1",
            },
        )
        if not rows:
            return {
                "provider": provider,
                "settings": {},
                "has_secret": False,
                "connection_status": "not_checked",
            }
        return rows[0]

    def save_provider_config(
        self,
        provider: str,
        settings: dict[str, Any],
        credentials: dict[str, Any] | None,
        updated_by: str,
    ) -> None:
        provider = str(provider or "").lower()
        if provider not in {"amocrm", "bitrix24"}:
            raise CrmConfigurationError("Неизвестный CRM-провайдер.")
        secret_value = json.dumps(credentials, ensure_ascii=False) if credentials else None
        self._request(
            "POST",
            "/rest/v1/rpc/admin_save_crm_provider_config",
            payload={
                "p_provider": provider,
                "p_settings": settings,
                "p_secret": secret_value,
                "p_updated_by": str(updated_by or "admin"),
            },
        )

    def get_provider_credentials(self, provider: str) -> dict[str, Any]:
        rows = self._request(
            "POST",
            "/rest/v1/rpc/admin_get_crm_provider_secret",
            payload={"p_provider": str(provider or "").lower()},
        )
        if not rows:
            return {}
        secret_value = rows[0].get("secret_value", "")
        try:
            value = json.loads(secret_value)
        except (TypeError, ValueError) as exc:
            raise CrmConfigurationError("Сохранённые CRM-данные повреждены.") from exc
        return value if isinstance(value, dict) else {}

    def set_connection_status(self, provider: str, check: ConnectionCheck) -> None:
        self._request(
            "PATCH",
            "/rest/v1/crm_provider_configs",
            params={"provider": f"eq.{str(provider or '').lower()}"},
            payload={
                "connection_status": "ok" if check.ok else "error",
                "connection_message": check.message[:500],
                "connection_checked_at": datetime.now(timezone.utc).isoformat(),
            },
            prefer="return=minimal",
        )

    def activate_provider(self, provider: str, updated_by: str) -> dict[str, Any]:
        result = self._request(
            "POST",
            "/rest/v1/rpc/admin_activate_crm_provider",
            payload={
                "p_provider": str(provider or "off").lower(),
                "p_updated_by": str(updated_by or "admin"),
            },
        )
        return normalize_runtime_settings(result or {})

    def get_delivery_logs(self, limit: int = 30) -> list[dict[str, Any]]:
        rows = self._request(
            "GET",
            "/rest/v1/crm_delivery_log",
            params={
                "select": "created_at,provider,event,status,message,lead_reference",
                "order": "created_at.desc",
                "limit": str(max(1, min(100, int(limit)))),
            },
        )
        return rows if isinstance(rows, list) else []


def create_store(secret_getter: Callable[[str, Any], Any]) -> SupabaseCrmStore:
    project_url = secret_getter("SUPABASE_URL", "")
    secret_key = secret_getter("SUPABASE_SERVICE_ROLE_KEY", "") or secret_getter(
        "SUPABASE_SECRET_KEY", ""
    )
    return SupabaseCrmStore(project_url, secret_key)


def normalize_amo_domain(value: Any) -> str:
    candidate = str(value or "").strip().lower()
    if not candidate:
        raise CrmConfigurationError("Укажите домен amoCRM.")
    if "://" not in candidate:
        candidate = f"https://{candidate}"
    parsed = urlparse(candidate)
    host = parsed.netloc.split("@")[ -1].split(":")[0]
    if parsed.scheme != "https" or not host or parsed.path not in {"", "/"}:
        raise CrmConfigurationError("Укажите только HTTPS-домен amoCRM без пути.")
    if not (host.endswith(".amocrm.ru") or host.endswith(".kommo.com")):
        raise CrmConfigurationError("Домен должен принадлежать amoCRM или Kommo.")
    return host


def test_amo_connection(
    settings: dict[str, Any],
    credentials: dict[str, Any],
    timeout: int = 15,
) -> ConnectionCheck:
    try:
        host = normalize_amo_domain(settings.get("domain"))
        token = str(credentials.get("access_token", "") or "").strip()
        if not token:
            raise CrmConfigurationError("Введите access token amoCRM.")
        response = requests.get(
            f"https://{host}/api/v4/account",
            headers={"Authorization": f"Bearer {token}"},
            timeout=timeout,
        )
        if response.status_code == 200:
            payload = response.json() if response.content else {}
            return ConnectionCheck(
                True,
                f"Подключение подтверждено: {payload.get('name') or host}",
                {"account_id": payload.get("id"), "account_name": payload.get("name")},
            )
        return ConnectionCheck(
            False,
            f"amoCRM вернула HTTP {response.status_code}. Проверьте домен и токен.",
            {"status_code": response.status_code},
        )
    except (requests.RequestException, ValueError, CrmConfigurationError) as exc:
        return ConnectionCheck(False, str(exc), {})
