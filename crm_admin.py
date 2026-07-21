from __future__ import annotations

import base64
import hashlib
import hmac
import time
from typing import Any, Callable

import streamlit as st

from crm_store import (
    CrmConfigurationError,
    DEFAULT_RUNTIME_SETTINGS,
    create_store,
    normalize_runtime_settings,
    test_amo_connection,
)


ADMIN_SESSION_TTL_SECONDS = 30 * 60


def is_admin_request() -> bool:
    try:
        return str(st.query_params.get("admin", "")).lower() in {"1", "true", "yes"}
    except Exception:
        return False


def _verify_admin_password(candidate: str, configured: str) -> bool:
    candidate = str(candidate or "")
    configured = str(configured or "")
    if configured.startswith("pbkdf2_sha256$"):
        try:
            _, iterations_text, salt_text, expected_text = configured.split("$", 3)
            iterations = int(iterations_text)
            salt = base64.urlsafe_b64decode(salt_text.encode("ascii"))
            expected = base64.urlsafe_b64decode(expected_text.encode("ascii"))
            actual = hashlib.pbkdf2_hmac(
                "sha256",
                candidate.encode("utf-8"),
                salt,
                iterations,
            )
            return hmac.compare_digest(actual, expected)
        except (TypeError, ValueError):
            return False
    if configured.startswith("sha256$"):
        expected = configured.split("$", 1)[1]
        actual = hashlib.sha256(candidate.encode("utf-8")).hexdigest()
        return hmac.compare_digest(actual, expected)
    return bool(configured) and hmac.compare_digest(candidate, configured)


def _admin_identity() -> str:
    return str(st.session_state.get("crm_admin_identity", "admin"))


def _render_login(secret_getter: Callable[[str, Any], Any]) -> bool:
    configured = str(secret_getter("ADMIN_PASSWORD_HASH", "") or "")
    if not configured:
        configured = str(secret_getter("ADMIN_PASSWORD", "") or "")
    if not configured:
        st.error("Доступ администратора ещё не настроен.")
        st.code('ADMIN_PASSWORD_HASH = "pbkdf2_sha256$..."', language="toml")
        st.caption("Добавьте хеш пароля в Secrets приложения Test.")
        return False

    authenticated_at = float(st.session_state.get("crm_admin_authenticated_at", 0) or 0)
    if st.session_state.get("crm_admin_authenticated") and time.time() - authenticated_at < ADMIN_SESSION_TTL_SECONDS:
        return True

    st.session_state.crm_admin_authenticated = False
    with st.form("crm_admin_login"):
        password = st.text_input("Пароль администратора", type="password")
        submitted = st.form_submit_button("Войти", type="primary", use_container_width=True)
    if submitted:
        if _verify_admin_password(password, configured):
            st.session_state.crm_admin_authenticated = True
            st.session_state.crm_admin_authenticated_at = time.time()
            st.session_state.crm_admin_identity = "admin"
            st.rerun()
        else:
            attempts = int(st.session_state.get("crm_admin_failed_attempts", 0) or 0) + 1
            st.session_state.crm_admin_failed_attempts = attempts
            st.error("Неверный пароль.")
    return False


@st.cache_data(ttl=60, show_spinner=False)
def _cached_runtime_settings(project_url: str, secret_key: str) -> dict[str, Any]:
    from crm_store import SupabaseCrmStore

    return SupabaseCrmStore(project_url, secret_key).get_runtime_settings()


def load_runtime_settings(secret_getter: Callable[[str, Any], Any]) -> dict[str, Any]:
    project_url = str(secret_getter("SUPABASE_URL", "") or "").strip()
    secret_key = str(
        secret_getter("SUPABASE_SERVICE_ROLE_KEY", "")
        or secret_getter("SUPABASE_SECRET_KEY", "")
        or ""
    ).strip()
    if not project_url or not secret_key:
        return dict(DEFAULT_RUNTIME_SETTINGS)
    try:
        return normalize_runtime_settings(_cached_runtime_settings(project_url, secret_key))
    except Exception:
        return dict(DEFAULT_RUNTIME_SETTINGS)


def _clear_runtime_cache() -> None:
    _cached_runtime_settings.clear()


def _render_storage_setup() -> None:
    st.warning("Постоянное хранилище админки ещё не подключено. CRM остаётся выключенной.")
    st.markdown(
        "1. Создайте проект Supabase.\n"
        "2. Выполните миграцию `db/001_crm_admin.sql`.\n"
        "3. Добавьте в Secrets приложения `SUPABASE_URL` и `SUPABASE_SERVICE_ROLE_KEY`."
    )


def _render_overview(store, runtime: dict[str, Any]) -> None:
    provider_labels = {"off": "Выключено", "amocrm": "amoCRM", "bitrix24": "Bitrix24"}
    format_labels = {"pptx": "Презентация", "xlsx": "Excel", "both": "Презентация + Excel"}
    col1, col2, col3 = st.columns(3)
    col1.metric("Активная CRM", provider_labels.get(runtime["active_provider"], "Выключено"))
    col2.metric("Формат заказчику", format_labels.get(runtime["customer_delivery_format"], "Презентация"))
    col3.metric("Режим", "Тестовый" if runtime.get("test_mode", True) else "Рабочий")

    st.subheader("Состояние подключений")
    for provider, label in (("amocrm", "amoCRM"), ("bitrix24", "Bitrix24")):
        config = store.get_provider_config(provider)
        status = config.get("connection_status", "not_checked")
        message = config.get("connection_message") or "Подключение не проверялось"
        icon = "✅" if status == "ok" else "⚪" if status == "not_checked" else "❌"
        st.write(f"{icon} **{label}:** {message}")


def _render_general_settings(store, runtime: dict[str, Any]) -> None:
    st.subheader("Выдача результата заказчику")
    with st.form("crm_runtime_settings"):
        customer_format = st.radio(
            "Формат заключения",
            options=["pptx", "xlsx", "both"],
            format_func=lambda value: {
                "pptx": "Презентация",
                "xlsx": "Excel",
                "both": "Оба файла",
            }[value],
            index={"pptx": 0, "xlsx": 1, "both": 2}.get(
                runtime.get("customer_delivery_format"), 0
            ),
            horizontal=True,
        )
        test_mode = st.toggle("Тестовый режим CRM", value=bool(runtime.get("test_mode", True)))
        saved = st.form_submit_button("Сохранить настройки", type="primary")
    if saved:
        runtime["customer_delivery_format"] = customer_format
        runtime["test_mode"] = test_mode
        store.save_runtime_settings(runtime, _admin_identity())
        _clear_runtime_cache()
        st.success("Настройки сохранены.")
        st.rerun()


def _render_telegram_settings(store, runtime: dict[str, Any]) -> None:
    st.subheader("Telegram")
    st.caption("Токен и chat ID остаются в Streamlit Secrets. Здесь настраивается состав отправки.")
    with st.form("telegram_runtime_settings"):
        diagnostics = st.toggle(
            "Отправлять техническую диагностику и ошибки",
            value=bool(runtime.get("telegram_diagnostics_enabled", True)),
        )
        lead_summary = st.toggle(
            "Отправлять текстовую сводку лида",
            value=bool(runtime.get("telegram_send_lead_summary", True)),
        )
        sales_playbook = st.toggle(
            "Прикладывать Sales Playbook",
            value=bool(runtime.get("telegram_send_sales_playbook", True)),
        )
        customer_report = st.toggle(
            "Прикладывать клиентское заключение",
            value=bool(runtime.get("telegram_send_customer_report", True)),
        )
        lead_template = st.text_area(
            "Шаблон сводки лида",
            value=str(runtime.get("telegram_lead_template", "")),
            placeholder="Оставьте пустым, чтобы использовать стандартную подробную сводку.",
            height=150,
            help="Переменные: {app}, {company}, {city}, {industry}, {email}, {phone}, {contact}, {role}, {score}, {ai}, {sales_digest}",
        )
        sales_caption = st.text_input(
            "Подпись к Sales Playbook",
            value=str(runtime.get("telegram_sales_caption", "[{app}] Sales playbook: {company}")),
        )
        customer_caption = st.text_input(
            "Подпись к клиентскому заключению",
            value=str(runtime.get("telegram_customer_caption", "[{app}] Клиентское заключение: {company}")),
        )
        saved = st.form_submit_button("Сохранить настройки Telegram", type="primary")
    if saved:
        runtime.update(
            {
                "telegram_diagnostics_enabled": diagnostics,
                "telegram_send_lead_summary": lead_summary,
                "telegram_send_sales_playbook": sales_playbook,
                "telegram_send_customer_report": customer_report,
                "telegram_lead_template": lead_template.strip(),
                "telegram_sales_caption": sales_caption.strip(),
                "telegram_customer_caption": customer_caption.strip(),
            }
        )
        store.save_runtime_settings(runtime, _admin_identity())
        _clear_runtime_cache()
        st.success("Настройки Telegram сохранены.")
        st.rerun()


def _render_amo_settings(store, runtime: dict[str, Any]) -> None:
    config = store.get_provider_config("amocrm")
    settings = config.get("settings") if isinstance(config.get("settings"), dict) else {}
    has_secret = bool(config.get("has_secret"))
    st.subheader("amoCRM")
    st.caption("Сохранённые токены не отображаются. Чтобы заменить токен, введите новый.")
    with st.form("amocrm_settings"):
        domain = st.text_input("Домен amoCRM", value=str(settings.get("domain", "")), placeholder="company.amocrm.ru")
        pipeline_id = st.text_input("ID воронки", value=str(settings.get("pipeline_id", "")))
        status_id = st.text_input("ID этапа", value=str(settings.get("status_id", "")))
        responsible_user_id = st.text_input(
            "ID ответственного", value=str(settings.get("responsible_user_id", ""))
        )
        task_due_hours = st.number_input(
            "Срок первой задачи, часов", min_value=1, max_value=720, value=int(settings.get("task_due_hours", 24) or 24)
        )
        access_token = st.text_input(
            "Новый access token" if has_secret else "Access token",
            type="password",
            placeholder="Оставьте пустым, чтобы сохранить текущий" if has_secret else "Введите токен",
        )
        save = st.form_submit_button("Сохранить конфигурацию", type="primary")

    new_settings = {
        "domain": domain.strip(),
        "pipeline_id": pipeline_id.strip(),
        "status_id": status_id.strip(),
        "responsible_user_id": responsible_user_id.strip(),
        "task_due_hours": int(task_due_hours),
    }
    if save:
        credentials = {"access_token": access_token.strip()} if access_token.strip() else None
        if not has_secret and not credentials:
            st.error("Для первого сохранения нужен access token.")
        else:
            store.save_provider_config("amocrm", new_settings, credentials, _admin_identity())
            st.success("Конфигурация amoCRM сохранена и ожидает проверки.")
            st.rerun()

    col_test, col_activate = st.columns(2)
    if col_test.button("Проверить подключение", use_container_width=True):
        credentials = {"access_token": access_token.strip()} if access_token.strip() else store.get_provider_credentials("amocrm")
        check = test_amo_connection(new_settings, credentials)
        store.set_connection_status("amocrm", check)
        if check.ok:
            st.success(check.message)
        else:
            st.error(check.message)
        st.rerun()

    can_activate = store.get_provider_config("amocrm").get("connection_status") == "ok"
    if col_activate.button(
        "Активировать amoCRM",
        disabled=not can_activate,
        type="primary",
        use_container_width=True,
    ):
        store.activate_provider("amocrm", _admin_identity())
        _clear_runtime_cache()
        st.success("amoCRM активирована для новых аудитов.")
        st.rerun()

    if runtime.get("active_provider") != "off" and st.button("Выключить отправку в CRM"):
        store.activate_provider("off", _admin_identity())
        _clear_runtime_cache()
        st.rerun()


def _render_bitrix_placeholder() -> None:
    st.subheader("Bitrix24")
    st.info("Адаптер Bitrix24 будет подключён после приёмки amoCRM на этапе X3-dev.4.")


def _render_logs(store) -> None:
    st.subheader("Журнал CRM")
    rows = store.get_delivery_logs(30)
    if not rows:
        st.caption("Событий пока нет.")
        return
    st.dataframe(rows, use_container_width=True, hide_index=True)


def render_crm_admin(app_version: str, secret_getter: Callable[[str, Any], Any]) -> None:
    st.title("CRM Control Center")
    st.caption(f"Khalil Audit System {app_version} | внутреннее администрирование")
    if not _render_login(secret_getter):
        return

    top_left, top_right = st.columns([4, 1])
    top_left.success("Защищённая административная сессия активна")
    if top_right.button("Выйти", use_container_width=True):
        st.session_state.crm_admin_authenticated = False
        st.session_state.crm_admin_authenticated_at = 0
        st.rerun()

    try:
        store = create_store(secret_getter)
        runtime = store.get_runtime_settings()
    except CrmConfigurationError:
        _render_storage_setup()
        return

    overview_tab, settings_tab, telegram_tab, crm_tab, logs_tab = st.tabs(
        ["Обзор", "Результат заказчику", "Telegram", "CRM", "Журнал"]
    )
    with overview_tab:
        _render_overview(store, runtime)
    with settings_tab:
        _render_general_settings(store, runtime)
    with telegram_tab:
        _render_telegram_settings(store, runtime)
    with crm_tab:
        provider = st.segmented_control(
            "Настраиваемая CRM",
            options=["amocrm", "bitrix24"],
            format_func=lambda value: "amoCRM" if value == "amocrm" else "Bitrix24",
            default="amocrm",
        )
        if provider == "bitrix24":
            _render_bitrix_placeholder()
        else:
            _render_amo_settings(store, runtime)
    with logs_tab:
        _render_logs(store)
