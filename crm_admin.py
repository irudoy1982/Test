from __future__ import annotations

import base64
import hashlib
import hmac
import re
import secrets
import time
from pathlib import Path
from typing import Any, Callable

import streamlit as st

from crm_store import (
    CrmConfigurationError,
    DEFAULT_RUNTIME_SETTINGS,
    create_store,
    normalize_runtime_settings,
    test_amo_connection,
)
from crm_assets import (
    validate_logo,
    validate_presentation_template,
    validate_vendor_matrix,
)


ADMIN_SESSION_TTL_SECONDS = 30 * 60
ADMIN_MAX_FAILED_ATTEMPTS = 5
ADMIN_LOCK_SECONDS = 60
ADMIN_PASSWORD_ITERATIONS = 600_000


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


def _hash_admin_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        ADMIN_PASSWORD_ITERATIONS,
    )
    salt_text = base64.urlsafe_b64encode(salt).decode("ascii")
    digest_text = base64.urlsafe_b64encode(digest).decode("ascii")
    return f"pbkdf2_sha256${ADMIN_PASSWORD_ITERATIONS}${salt_text}${digest_text}"


def _admin_identity() -> str:
    return str(st.session_state.get("crm_admin_identity", "admin"))


def _admin_role() -> str:
    return str(st.session_state.get("crm_admin_role", "viewer"))


def _render_login(secret_getter: Callable[[str, Any], Any], store=None) -> bool:
    configured_username = str(secret_getter("ADMIN_USERNAME", "admin") or "admin").strip()
    configured = str(secret_getter("ADMIN_PASSWORD_HASH", "") or "")
    database_auth_available = False
    if store is not None:
        try:
            database_auth_available = bool(store.list_admin_users())
        except Exception:
            database_auth_available = False
    if not configured and not database_auth_available:
        configured = str(secret_getter("ADMIN_PASSWORD", "") or "")
    if not configured:
        st.error("Доступ администратора ещё не настроен.")
        st.code(
            'ADMIN_USERNAME = "admin"\nADMIN_PASSWORD_HASH = "pbkdf2_sha256$..."',
            language="toml",
        )
        st.caption("Добавьте логин и хеш пароля в Secrets приложения Test.")
        return False

    authenticated_at = float(st.session_state.get("crm_admin_authenticated_at", 0) or 0)
    if st.session_state.get("crm_admin_authenticated") and time.time() - authenticated_at < ADMIN_SESSION_TTL_SECONDS:
        return True

    st.session_state.crm_admin_authenticated = False
    locked_until = float(st.session_state.get("crm_admin_locked_until", 0) or 0)
    if locked_until > time.time():
        wait_seconds = max(1, int(locked_until - time.time()))
        st.error(f"Слишком много неудачных попыток. Повторите вход через {wait_seconds} сек.")
        return False

    login_col, _ = st.columns([1, 1])
    with login_col:
        with st.form("crm_admin_login"):
            username = st.text_input("Логин", autocomplete="username")
            password = st.text_input(
                "Пароль",
                type="password",
                autocomplete="current-password",
            )
            submitted = st.form_submit_button("Войти", type="primary", use_container_width=True)
    if submitted:
        identity = username.strip()
        authenticated_role = ""
        authenticated_name = identity
        username_ok = hmac.compare_digest(identity, configured_username)
        if configured and username_ok and _verify_admin_password(password, configured):
            authenticated_role = "admin"
            authenticated_name = configured_username
        elif store is not None:
            try:
                account = store.get_admin_user(identity)
            except Exception:
                account = {}
            if (
                account.get("active")
                and _verify_admin_password(password, str(account.get("password_hash") or ""))
            ):
                authenticated_role = str(account.get("role") or "viewer")
                authenticated_name = str(account.get("display_name") or identity)

        if authenticated_role:
            st.session_state.crm_admin_authenticated = True
            st.session_state.crm_admin_authenticated_at = time.time()
            st.session_state.crm_admin_identity = identity
            st.session_state.crm_admin_display_name = authenticated_name
            st.session_state.crm_admin_role = authenticated_role
            st.session_state.crm_admin_failed_attempts = 0
            st.session_state.crm_admin_locked_until = 0
            st.rerun()
        else:
            attempts = int(st.session_state.get("crm_admin_failed_attempts", 0) or 0) + 1
            st.session_state.crm_admin_failed_attempts = attempts
            if attempts >= ADMIN_MAX_FAILED_ATTEMPTS:
                st.session_state.crm_admin_locked_until = time.time() + ADMIN_LOCK_SECONDS
                st.session_state.crm_admin_failed_attempts = 0
                st.error("Вход временно заблокирован после нескольких неудачных попыток.")
            else:
                st.error("Неверный логин или пароль.")
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


@st.cache_data(ttl=300, show_spinner=False)
def _cached_runtime_asset(project_url: str, secret_key: str, asset_key: str) -> bytes | None:
    from crm_store import SupabaseCrmStore

    return SupabaseCrmStore(project_url, secret_key).download_asset(asset_key)


def load_runtime_asset_bytes(
    secret_getter: Callable[[str, Any], Any],
    asset_key: str,
) -> bytes | None:
    project_url = str(secret_getter("SUPABASE_URL", "") or "").strip()
    secret_key = str(
        secret_getter("SUPABASE_SERVICE_ROLE_KEY", "")
        or secret_getter("SUPABASE_SECRET_KEY", "")
        or ""
    ).strip()
    if not project_url or not secret_key:
        return None
    try:
        return _cached_runtime_asset(project_url, secret_key, asset_key)
    except Exception:
        return None


def _clear_runtime_cache() -> None:
    _cached_runtime_settings.clear()


def _clear_asset_cache() -> None:
    _cached_runtime_asset.clear()


def _render_storage_setup() -> None:
    st.warning("Постоянное хранилище админки ещё не подключено. CRM остаётся выключенной.")
    st.markdown(
        "1. Создайте проект Supabase.\n"
        "2. Выполните по порядку миграции `db/001_crm_admin.sql`, "
        "`db/002_admin_assets.sql`, `db/003_admin_users.sql`.\n"
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


def _fallback_asset(asset_key: str) -> tuple[str, bytes, str]:
    root = Path(__file__).resolve().parent
    fallback = {
        "logo": (root / "logo.png", "image/png"),
        "presentation_template": (
            root / "static" / "audit_presentation_khalil.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ),
        "vendor_matrix": (
            root / "vendor_matrix_detailed.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ),
    }[asset_key]
    path, content_type = fallback
    return path.name, path.read_bytes(), content_type


def _render_asset_editor(
    store,
    *,
    asset_key: str,
    title: str,
    description: str,
    allowed_types: list[str],
    validator,
) -> None:
    st.markdown(f"#### {title}")
    st.caption(description)
    metadata = store.get_asset_metadata(asset_key)
    remote_data = store.download_asset(asset_key)
    if remote_data:
        filename = str(metadata.get("filename") or asset_key)
        content_type = str(metadata.get("content_type") or "application/octet-stream")
        active_data = remote_data
        source_label = "Опубликованная версия из админки"
    else:
        filename, active_data, content_type = _fallback_asset(asset_key)
        source_label = "Стабильная версия из репозитория"

    info_col, download_col = st.columns([3, 1])
    info_col.write(f"**Активный файл:** {filename}")
    info_col.caption(
        f"{source_label} · {len(active_data) / 1024:.1f} КБ"
        + (f" · обновлён {metadata.get('updated_at')}" if metadata.get("updated_at") else "")
    )
    download_col.download_button(
        "Скачать образец",
        data=active_data,
        file_name=filename,
        mime=content_type,
        key=f"admin_asset_download_{asset_key}",
        use_container_width=True,
    )

    uploaded = st.file_uploader(
        "Выберите новую версию",
        type=allowed_types,
        key=f"admin_asset_upload_{asset_key}",
    )
    if uploaded is not None:
        candidate = uploaded.getvalue()
        validation = validator(candidate, uploaded.name)
        if validation.ok:
            st.success(validation.message)
        else:
            st.error(validation.message)
        if st.button(
            "Проверить и опубликовать",
            key=f"admin_asset_publish_{asset_key}",
            type="primary",
            disabled=not validation.ok,
        ):
            store.save_asset(
                asset_key,
                uploaded.name,
                validation.content_type,
                candidate,
                validation.details,
                _admin_identity(),
            )
            _clear_asset_cache()
            st.success("Новая версия опубликована. Предыдущая сохранена в истории.")
            st.rerun()


def _render_assets(store) -> None:
    st.subheader("Бренд и шаблоны")
    st.caption(
        "Файлы начинают использоваться приложением после публикации. "
        "Если внешний файл недоступен, Test автоматически вернётся к версии из репозитория."
    )
    _render_asset_editor(
        store,
        asset_key="logo",
        title="Логотип веб-анкеты",
        description="PNG или JPEG до 5 МБ. Логотип внутри презентации меняется вместе с PPTX-шаблоном.",
        allowed_types=["png", "jpg", "jpeg"],
        validator=validate_logo,
    )
    st.divider()
    _render_asset_editor(
        store,
        asset_key="presentation_template",
        title="Шаблон презентации",
        description="Используйте скачанный образец: служебные поля в двойных фигурных скобках нельзя удалять.",
        allowed_types=["pptx"],
        validator=validate_presentation_template,
    )
    st.divider()
    _render_asset_editor(
        store,
        asset_key="vendor_matrix",
        title="Портфель производителей и решений",
        description="Скачайте действующий формат, измените строки и загрузите XLSX обратно.",
        allowed_types=["xlsx"],
        validator=validate_vendor_matrix,
    )


def _valid_admin_username(value: str) -> bool:
    return bool(re.fullmatch(r"[a-zA-Z0-9._@-]{3,80}", str(value or "").strip()))


def _render_users(store) -> None:
    st.subheader("Пользователи админки")
    st.caption("Пароли хранятся только в виде стойких хешей. Пользователь не может отключить сам себя.")
    users = store.list_admin_users()
    role_labels = {"admin": "Администратор", "editor": "Редактор", "viewer": "Наблюдатель"}
    if users:
        visible_rows = [
            {
                "Логин": row.get("username"),
                "Имя": row.get("display_name"),
                "Роль": role_labels.get(row.get("role"), row.get("role")),
                "Активен": bool(row.get("active")),
                "Изменён": row.get("updated_at"),
            }
            for row in users
        ]
        st.dataframe(visible_rows, use_container_width=True, hide_index=True)
    else:
        st.info("Дополнительных пользователей пока нет. Начальный администратор берётся из Secrets.")

    with st.expander("Создать пользователя", expanded=not users):
        with st.form("admin_user_create"):
            username = st.text_input("Логин нового пользователя")
            display_name = st.text_input("Имя")
            role = st.selectbox(
                "Роль",
                options=["editor", "viewer", "admin"],
                format_func=lambda value: role_labels[value],
            )
            password = st.text_input("Временный пароль", type="password")
            confirmation = st.text_input("Повторите пароль", type="password")
            create_user = st.form_submit_button("Создать пользователя", type="primary")
        if create_user:
            if not _valid_admin_username(username):
                st.error("Логин: 3-80 символов, только латиница, цифры и . _ @ -")
            elif len(password) < 12:
                st.error("Пароль должен содержать не менее 12 символов.")
            elif password != confirmation:
                st.error("Пароли не совпадают.")
            elif store.get_admin_user(username.strip()):
                st.error("Пользователь с таким логином уже существует.")
            else:
                store.save_admin_user(
                    username.strip(),
                    display_name.strip() or username.strip(),
                    role,
                    _hash_admin_password(password),
                    _admin_identity(),
                )
                st.success("Пользователь создан.")
                st.rerun()

    if users:
        st.markdown("#### Изменить пользователя")
        selected_username = st.selectbox(
            "Пользователь",
            options=[str(row.get("username")) for row in users],
            key="admin_user_selected",
        )
        selected = next(row for row in users if row.get("username") == selected_username)
        with st.form("admin_user_edit"):
            edit_name = st.text_input("Имя", value=str(selected.get("display_name") or ""))
            edit_role = st.selectbox(
                "Роль",
                options=["admin", "editor", "viewer"],
                index=["admin", "editor", "viewer"].index(str(selected.get("role") or "viewer")),
                format_func=lambda value: role_labels[value],
            )
            active = st.toggle("Доступ активен", value=bool(selected.get("active", True)))
            new_password = st.text_input("Новый пароль (необязательно)", type="password")
            save_user = st.form_submit_button("Сохранить изменения", type="primary")
        if save_user:
            if selected_username == _admin_identity() and not active:
                st.error("Нельзя отключить собственную учётную запись.")
            elif new_password and len(new_password) < 12:
                st.error("Новый пароль должен содержать не менее 12 символов.")
            else:
                store.save_admin_user(
                    selected_username,
                    edit_name.strip() or selected_username,
                    edit_role,
                    _hash_admin_password(new_password) if new_password else None,
                    _admin_identity(),
                )
                store.set_admin_user_active(
                    selected_username,
                    active,
                    _admin_identity(),
                )
                st.success("Пользователь обновлён.")
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
    try:
        store = create_store(secret_getter)
    except CrmConfigurationError:
        store = None
    if not _render_login(secret_getter, store):
        return

    top_left, top_right = st.columns([4, 1])
    role_labels = {"admin": "администратор", "editor": "редактор", "viewer": "наблюдатель"}
    display_name = str(st.session_state.get("crm_admin_display_name", _admin_identity()))
    top_left.success(
        f"Выполнен вход: {display_name} · {role_labels.get(_admin_role(), _admin_role())}"
    )
    if top_right.button("Выйти", use_container_width=True):
        st.session_state.crm_admin_authenticated = False
        st.session_state.crm_admin_authenticated_at = 0
        st.session_state.crm_admin_role = ""
        st.rerun()

    if store is None:
        _render_storage_setup()
        return
    runtime = store.get_runtime_settings()

    if _admin_role() == "viewer":
        overview_tab, logs_tab = st.tabs(["Обзор", "Журнал"])
        with overview_tab:
            _render_overview(store, runtime)
        with logs_tab:
            _render_logs(store)
        return

    tab_names = ["Обзор", "Результат заказчику", "Telegram", "Бренд и шаблоны", "CRM", "Журнал"]
    if _admin_role() == "admin":
        tab_names.insert(5, "Пользователи")
    tabs = st.tabs(tab_names)
    tab_index = 0
    with tabs[tab_index]:
        _render_overview(store, runtime)
    tab_index += 1
    with tabs[tab_index]:
        _render_general_settings(store, runtime)
    tab_index += 1
    with tabs[tab_index]:
        _render_telegram_settings(store, runtime)
    tab_index += 1
    with tabs[tab_index]:
        try:
            _render_assets(store)
        except CrmConfigurationError as exc:
            st.error(f"Раздел файлов не готов: {exc}")
            st.caption("Выполните миграцию db/002_admin_assets.sql в Supabase.")
    tab_index += 1
    with tabs[tab_index]:
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
    tab_index += 1
    if _admin_role() == "admin":
        with tabs[tab_index]:
            try:
                _render_users(store)
            except CrmConfigurationError as exc:
                st.error(f"Управление пользователями не готово: {exc}")
                st.caption("Выполните миграцию db/003_admin_users.sql в Supabase.")
        tab_index += 1
    with tabs[tab_index]:
        _render_logs(store)
