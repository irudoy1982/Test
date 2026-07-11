import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import html
import base64
import json
import re
import zlib
import threading
import time
import random
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

REQUEST_VERIFY = True


def configure_https_certificates():
    global REQUEST_VERIFY
    try:
        import certifi
        cert_path = certifi.where()
        os.environ.setdefault("SSL_CERT_FILE", cert_path)
        os.environ.setdefault("REQUESTS_CA_BUNDLE", cert_path)
        os.environ.setdefault("GRPC_DEFAULT_SSL_ROOTS_FILE_PATH", cert_path)
        REQUEST_VERIFY = cert_path
    except Exception:
        REQUEST_VERIFY = True


configure_https_certificates()

#----------ИИ-----------
# --- AI BLOCK START ---

def sanitize_for_ai(c_info, results):
    forbidden = [
        "Наименование компании",
        "Сайт компании",
        "Email",
        "ФИО контактного лица",
        "Должность",
        "Контактный телефон",
        "Город",
        "Компания",
        "Контакт",
        "Телефон",
        "client_",
        "name_",
        "email",
        "phone",
    ]
    safe_client = {
        "Сфера деятельности": c_info.get("Сфера деятельности", "")
    }
    safe_results = {
        k: v for k, v in results.items()
        if not any(f.lower() in str(k).lower() for f in forbidden)
    }
    return safe_client, safe_results

def load_vendor_matrix():
    try:
        df = pd.read_excel("Портфель для отчета.xlsx")

        vendors_text = ""

        for _, row in df.iterrows():
            row_text = " | ".join(
                [str(x) for x in row.values if pd.notna(x)]
            )
            vendors_text += row_text + "\n"

        return vendors_text

    except Exception as e:
        return f"Ошибка загрузки вендоров: {e}"


def load_vendor_names():
    try:
        detailed = load_detailed_vendor_names()
        if detailed:
            return detailed
        df = pd.read_excel("Портфель для отчета.xlsx", header=None)
        values = []
        for value in df.iloc[:, 0].dropna().tolist():
            vendor = str(value).strip()
            if vendor and vendor.lower() not in {"nan", "none"}:
                values.append(vendor)
        return list(dict.fromkeys(values))
    except Exception:
        return []


def get_app_secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


APP_INSTANCE_DEFAULT = "Test"


def get_app_instance_label():
    value = str(get_app_secret("APP_INSTANCE", APP_INSTANCE_DEFAULT) or "").strip()
    return value or APP_INSTANCE_DEFAULT


def redact_secret(text, *secrets):
    safe_text = str(text)
    for secret in secrets:
        if secret:
            safe_text = safe_text.replace(str(secret), "***")
    return safe_text


def find_node_exe():
    import shutil

    node_exe = get_app_secret("NODE_EXE", None)
    if node_exe:
        return node_exe

    node_exe = shutil.which("node")
    if node_exe:
        return node_exe

    local_node = os.path.expanduser(
        r"~\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
    )
    if os.path.exists(local_node):
        return local_node

    return None


def node_fetch_json(url, payload, timeout_seconds=25, env_extra=None):
    import subprocess
    import tempfile

    node_exe = find_node_exe()
    if not node_exe:
        raise RuntimeError("Node.js не найден для HTTPS-запроса.")

    node_script = r"""
const fs = require('fs');
const requestPath = process.argv[2];
const request = JSON.parse(fs.readFileSync(requestPath, 'utf8'));
const controller = new AbortController();
const timer = setTimeout(() => controller.abort(), request.timeoutMs || 25000);

(async () => {
  const headers = request.headers || {};
  let body;

  if (request.multipart) {
    body = new FormData();
    for (const [key, value] of Object.entries(request.fields || {})) {
      body.append(key, String(value));
    }
    for (const file of request.files || []) {
      const bytes = await fs.promises.readFile(file.path);
      body.append(file.field, new Blob([bytes]), file.filename);
    }
  } else {
    headers['Content-Type'] = 'application/json';
    body = JSON.stringify(request.body || {});
  }

  const response = await fetch(request.url, {
    method: request.method || 'POST',
    headers,
    body,
    signal: controller.signal,
  });
  const text = await response.text();
  clearTimeout(timer);

  if (!response.ok) {
    console.error(text.slice(0, 1500));
    process.exit(2);
  }

  process.stdout.write(text || '{}');
})().catch((error) => {
  clearTimeout(timer);
  console.error(`${error.name}: ${error.message}`);
  process.exit(1);
});
"""
    temp_paths = []
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as request_file:
            json.dump(payload, request_file, ensure_ascii=False)
            request_path = request_file.name
            temp_paths.append(request_path)
        with tempfile.NamedTemporaryFile("w", suffix=".cjs", delete=False, encoding="utf-8") as script_file:
            script_file.write(node_script)
            script_path = script_file.name
            temp_paths.append(script_path)

        env = os.environ.copy()
        env["NODE_TLS_REJECT_UNAUTHORIZED"] = "0"
        env["NODE_OPTIONS"] = f"{env.get('NODE_OPTIONS', '')} --no-warnings".strip()
        if env_extra:
            env.update(env_extra)

        completed = subprocess.run(
            [node_exe, script_path, request_path],
            capture_output=True,
            text=True,
            timeout=timeout_seconds + 5,
            env=env,
        )
        if completed.returncode != 0:
            stderr_lines = []
            for line in completed.stderr.splitlines():
                if "Assertion failed:" in line or "UV_HANDLE_CLOSING" in line:
                    continue
                if "NODE_TLS_REJECT_UNAUTHORIZED" in line:
                    continue
                stderr_lines.append(line)
            stderr_text = "\n".join(stderr_lines).strip()
            if len(stderr_text) > 1200:
                stderr_text = stderr_text[:1200] + "..."
            raise RuntimeError(stderr_text or "Node HTTPS request failed")
        return json.loads(completed.stdout or "{}")
    finally:
        for path in temp_paths:
            try:
                os.unlink(path)
            except OSError:
                pass


def telegram_send_node(token, method, fields, files=None, timeout_seconds=10):
    import requests

    upload_files = {}
    try:
        for file_item in files or []:
            field_name = file_item.get("field", "document")
            upload_files[field_name] = (
                file_item["filename"],
                BytesIO(file_item["bytes"]),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        response = requests.post(
            f"https://api.telegram.org/bot{token}/{method}",
            data=fields,
            files=upload_files or None,
            timeout=timeout_seconds,
            verify=REQUEST_VERIFY,
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.SSLError:
        response = requests.post(
            f"https://api.telegram.org/bot{token}/{method}",
            data=fields,
            files=upload_files or None,
            timeout=timeout_seconds,
            verify=False,
        )
        response.raise_for_status()
        return response.json()


def build_telegram_lead_text(client_info, final_score, sales_digest):
    return (
        "🚨 Новый запрос на аудит!\n"
        f"📌 Приложение: {get_app_instance_label()}\n"
        f"🏢 Компания: {client_info.get('Наименование компании', '-')}\n"
        f"📍 Город: {client_info.get('Город', '-')}\n"
        f"📊 Сфера: {client_info.get('Сфера деятельности', '-')}\n"
        f"🌐 Сайт: {client_info.get('Сайт компании', '-')}\n"
        f"📧 Email: {client_info.get('Email', '-')}\n"
        f"📞 Телефон: {client_info.get('Контактный телефон', '-')}\n"
        f"👤 Контакт: {client_info.get('ФИО контактного лица', '-')}\n"
        f"💼 Должность: {client_info.get('Должность', '-')}\n"
        f"📊 Уровень зрелости: {final_score}%\n\n"
        f"💡 Что предложить первым:\n{sales_digest}"
    )


def build_telegram_ai_failure_text(client_info, final_score, ai_error):
    safe_error = str(ai_error or "Неизвестная ошибка").strip()
    if len(safe_error) > 1800:
        safe_error = safe_error[:1800] + "..."

    return (
        "ℹ️ Gemini временно недоступен; отчет собран экспертным движком Khalil Audit\n"
        f"📌 Приложение: {get_app_instance_label()}\n"
        f"🏢 Компания: {client_info.get('Наименование компании', '-')}\n"
        f"📍 Город: {client_info.get('Город', '-')}\n"
        f"📊 Сфера: {client_info.get('Сфера деятельности', '-')}\n"
        f"🌐 Сайт: {client_info.get('Сайт компании', '-')}\n"
        f"📧 Email: {client_info.get('Email', '-')}\n"
        f"📞 Телефон: {client_info.get('Контактный телефон', '-')}\n"
        f"👤 Контакт: {client_info.get('ФИО контактного лица', '-')}\n"
        f"💼 Должность: {client_info.get('Должность', '-')}\n"
        f"📊 Уровень зрелости: {final_score}%\n\n"
        f"Внутренняя диагностика: {safe_error}"
    )


def build_telegram_generation_started_text(client_info, final_score):
    return (
        "🟢 Начато формирование аудита Khalil Audit\n"
        f"📌 Приложение: {get_app_instance_label()}\n"
        f"🏢 Компания: {client_info.get('Наименование компании', '-')}\n"
        f"📍 Город: {client_info.get('Город', '-')}\n"
        f"📊 Сфера: {client_info.get('Сфера деятельности', '-')}\n"
        f"🌐 Сайт: {client_info.get('Сайт компании', '-')}\n"
        f"📧 Email: {client_info.get('Email', '-')}\n"
        f"📞 Телефон: {client_info.get('Контактный телефон', '-')}\n"
        f"👤 Контакт: {client_info.get('ФИО контактного лица', '-')}\n"
        f"💼 Должность: {client_info.get('Должность', '-')}\n"
        f"📊 Предварительная зрелость: {final_score}%"
    )


def build_telegram_generation_error_text(client_info, final_score, error):
    safe_error = str(error or "Неизвестная ошибка").strip()
    if len(safe_error) > 1800:
        safe_error = safe_error[:1800] + "..."

    return (
        "🔴 Ошибка формирования аудита Khalil Audit\n"
        f"📌 Приложение: {get_app_instance_label()}\n"
        f"🏢 Компания: {client_info.get('Наименование компании', '-')}\n"
        f"📍 Город: {client_info.get('Город', '-')}\n"
        f"📧 Email: {client_info.get('Email', '-')}\n"
        f"📞 Телефон: {client_info.get('Контактный телефон', '-')}\n"
        f"👤 Контакт: {client_info.get('ФИО контактного лица', '-')}\n"
        f"📊 Предварительная зрелость: {final_score}%\n\n"
        f"Диагностика: {safe_error}"
    )


def send_internal_telegram_message(text, timeout_seconds=8):
    if not TOKEN or not CHAT_ID:
        return "Telegram не отправлен: не найдены TELEGRAM_TOKEN или TELEGRAM_CHAT_ID."

    try:
        telegram_send_node(
            TOKEN,
            "sendMessage",
            {"chat_id": CHAT_ID, "text": text},
            timeout_seconds=timeout_seconds
        )
        return "ok"
    except Exception as exc:
        return f"Telegram не отправлен: {redact_secret(exc, TOKEN)}"


def normalize_ai_risks_payload(payload):
    def repair_mojibake(value):
        text = str(value)
        suspicious = sum(text.count(marker) for marker in ("Р", "С", "Ð", "Ñ"))
        if suspicious < 3:
            return text

        try:
            repaired = text.encode("cp1251").decode("utf-8")
        except UnicodeError:
            return text

        original_badness = text.count("Р") + text.count("С") + text.count("Ð") + text.count("Ñ")
        repaired_badness = repaired.count("Р") + repaired.count("С") + repaired.count("Ð") + repaired.count("Ñ")
        return repaired if repaired_badness < original_badness else text

    def normalize_risk_item(item):
        if not isinstance(item, dict):
            return None

        risk = repair_mojibake(item.get("risk", "")).strip()
        recommendation = repair_mojibake(item.get("recommendation", "")).strip()
        if not risk or not recommendation:
            return None

        vendors = item.get("vendors", [])
        regulators = item.get("regulators", [])
        if not isinstance(vendors, list):
            vendors = [vendors] if vendors else []
        if not isinstance(regulators, list):
            regulators = [regulators] if regulators else []

        return {
            "level": repair_mojibake(item.get("level", "MEDIUM")).strip() or "MEDIUM",
            "risk": risk,
            "description": repair_mojibake(item.get("description", risk)).strip() or risk,
            "impact": repair_mojibake(item.get("impact", "Риск может привести к снижению устойчивости ИТ/ИБ процессов.")).strip(),
            "recommendation": recommendation,
            "vendors": [repair_mojibake(value).strip() for value in vendors if repair_mojibake(value).strip()][:3],
            "regulators": [repair_mojibake(value).strip() for value in regulators if repair_mojibake(value).strip()][:3],
        }

    def collect_from_list(value, source_area="ИИ"):
        normalized_items = []
        if not isinstance(value, list):
            return normalized_items
        for item in value:
            normalized = normalize_risk_item(item)
            if normalized:
                normalized["_ai_area"] = source_area
                normalized_items.append(normalized)
        return normalized_items

    if isinstance(payload, list):
        return [
            normalized
            for normalized in collect_from_list(payload)
            if normalized
        ]

    if isinstance(payload, dict):
        combined = []
        for key, source_area in (
            ("it_recommendations", "ИТ"),
            ("security_recommendations", "ИБ"),
            ("risks", "ИИ"),
            ("recommendations", "ИИ"),
            ("items", "ИИ"),
            ("data", "ИИ"),
        ):
            combined.extend(collect_from_list(payload.get(key), source_area))
        if combined:
            return combined

        for key in ("risks", "recommendations", "items", "data"):
            value = payload.get(key)
            if isinstance(value, list):
                return [
                    normalized
                    for normalized in (normalize_risk_item(item) for item in value)
                    if normalized
                ]

    return []


def normalize_ai_audit_narrative(payload):
    if not isinstance(payload, dict):
        return {}

    def clean_text(value):
        text = str(value or "").strip()
        suspicious = sum(text.count(marker) for marker in ("Р", "С", "Ð", "Ñ"))
        if suspicious >= 3:
            try:
                repaired = text.encode("cp1251").decode("utf-8")
                if repaired.count("Р") + repaired.count("С") < text.count("Р") + text.count("С"):
                    text = repaired
            except UnicodeError:
                pass
        return re.sub(r"\s+", " ", text).strip()

    def clean_list(values, limit=8):
        if not isinstance(values, list):
            return []
        cleaned = []
        for value in values:
            if isinstance(value, dict):
                text = clean_text(value.get("text") or value.get("summary") or value.get("recommendation"))
            else:
                text = clean_text(value)
            if text:
                cleaned.append(text)
            if len(cleaned) >= limit:
                break
        return cleaned

    def clean_observations(values, limit=6):
        if not isinstance(values, list):
            return []
        cleaned = []
        for value in values:
            if isinstance(value, dict):
                title = clean_text(value.get("title") or value.get("domain") or value.get("risk"))
                text = clean_text(value.get("text") or value.get("description") or value.get("recommendation"))
            else:
                title = "Наблюдение"
                text = clean_text(value)
            if title and text:
                cleaned.append((title[:80], text))
            if len(cleaned) >= limit:
                break
        return cleaned

    def clean_roadmap(values, limit=8):
        if not isinstance(values, list):
            return []
        rows = []
        for value in values:
            if not isinstance(value, dict):
                continue
            action = clean_text(value.get("action") or value.get("recommendation"))
            rationale = clean_text(value.get("rationale") or value.get("why") or value.get("impact"))
            if not action:
                continue
            rows.append({
                "phase": clean_text(value.get("phase") or "0-90 дней"),
                "priority": clean_text(value.get("priority") or "P2"),
                "domain": clean_text(value.get("domain") or "ИТ/ИБ"),
                "action": action,
                "rationale": rationale or "Мера снижает выявленный риск и повышает управляемость.",
                "owner": clean_text(value.get("owner") or "ИТ/ИБ"),
                "effort": clean_text(value.get("effort") or "Средняя"),
            })
            if len(rows) >= limit:
                break
        return rows

    return {
        "executive_summary": clean_list(payload.get("executive_summary"), limit=7),
        "audit_observations": clean_observations(payload.get("audit_observations"), limit=6),
        "management_decisions": clean_list(payload.get("management_decisions"), limit=6),
        "roadmap": clean_roadmap(payload.get("roadmap"), limit=8),
    }


def is_truncated_ai_text(value):
    text = str(value or "").strip()
    if not text:
        return True
    if len(text) < 20:
        return True
    unfinished_endings = (
        " и", " в", " с", " на", " для", " по", " от", " до", " при",
        "1.", "2.", "3.", " 1", " 2", " 3"
    )
    return text.endswith(unfinished_endings)


def prepare_ai_risks_for_report(items, min_items=1):
    if not isinstance(items, list):
        return []

    prepared = []
    seen = set()
    for item in items:
        if not isinstance(item, dict):
            continue
        if is_truncated_ai_text(item.get("risk")):
            continue
        if is_truncated_ai_text(item.get("recommendation")):
            continue

        semantic_key = risk_semantic_key(item)
        if not semantic_key or semantic_key in seen:
            continue
        prepared.append(item)
        seen.add(semantic_key)

    return prepared if len(prepared) >= min_items else []


def ai_quality_gate(items, min_items=6):
    prepared = prepare_ai_risks_for_report(items, min_items=1)
    if len(prepared) < min_items:
        return [], f"ИИ дал только {len(prepared)} пригодных пунктов из минимально ожидаемых {min_items}."

    security_markers = (
        "mfa", "edr", "xdr", "mdr", "epp", "siem", "soc", "pam", "dlp",
        "waf", "ids", "ips", "ztna", "mail", "почт", "уязв", "patch",
        "учет", "доступ", "endpoint", "мониторинг событий", "реагирован"
    )
    it_markers = (
        "backup", "резерв", "dr", "rto", "rpo", "сервер", "сеть", "схд",
        "виртуал", "мониторинг", "обновлен", "ос", "инфраструкт",
        "эксплуатац", "емкост", "производительност", "бизнес-систем"
    )

    security_count = 0
    it_count = 0
    weak_text_count = 0
    for item in prepared:
        combined = " ".join(
            str(item.get(field, ""))
            for field in ("risk", "description", "impact", "recommendation")
        ).lower()
        if any(marker in combined for marker in security_markers):
            security_count += 1
        if any(marker in combined for marker in it_markers):
            it_count += 1
        if len(str(item.get("recommendation", "")).strip()) < 90:
            weak_text_count += 1

    if weak_text_count > max(2, len(prepared) // 3):
        return [], "ИИ дал слишком короткие рекомендации; включены экспертные правила."
    if security_count < 3:
        return [], "ИИ почти не покрыл ИБ-домены; включены экспертные правила."
    if it_count < 3:
        return [], "ИИ почти не покрыл ИТ-инфраструктуру; включены экспертные правила."

    return prepared, ""


def security_control_snapshot(results):
    controls = [
        "MFA",
        "EPP",
        "EDR",
        "XDR",
        "MDR",
        "DLP",
        "Mail Security",
        "WAF",
        "Anti-DDoS",
        "IDS/IPS",
        "NAC",
        "ZTNA",
        "IAM",
        "PAM",
        "SIEM",
        "SOAR",
        "NAD",
        "Patch Management",
        "SAST",
        "DAST",
    ]
    enabled = []
    missing = []
    for control in controls:
        value = results.get(control)
        if is_enabled(value):
            enabled.append(f"{control}: {value}")
        else:
            missing.append(control)
    return enabled, missing


def risk_conflicts_with_answers(item, results):
    key = risk_semantic_key(item)

    if key == "mfa" and is_enabled(results.get("MFA")):
        return "MFA already enabled"
    if key == "pam" and is_enabled(results.get("PAM")):
        return "PAM already enabled"
    if key == "siem_soc" and is_enabled(results.get("SIEM")):
        return "SIEM already enabled"
    if key == "patch" and is_enabled(results.get("Patch Management")):
        return "Patch Management already enabled"
    if key == "web_waf" and is_enabled(results.get("WAF")):
        return "WAF already enabled"
    if key == "mail" and is_enabled(results.get("Mail Security")):
        return "Mail Security already enabled"
    if key == "dlp" and is_enabled(results.get("DLP")):
        return "DLP already enabled"
    if key == "backup" and is_enabled(results.get("Резервное копирование")):
        return "Backup already enabled"
    if key == "endpoint_detection" and any(
        is_enabled(results.get(control))
        for control in ("EDR", "XDR", "MDR")
    ):
        return "Endpoint detection/response already enabled"

    return ""


def filter_ai_risks_by_answers(items, results):
    filtered = []
    skipped = []
    for item in items:
        conflict = risk_conflicts_with_answers(item, results)
        if conflict:
            skipped.append(f"{risk_semantic_key(item)}: {conflict}")
            continue
        filtered.append(item)
    return filtered, skipped


def normalize_site_domain(site):
    value = str(site or "").strip().lower()
    value = re.sub(r"^https?://", "", value)
    value = re.sub(r"^www\.", "", value)
    value = value.split("/")[0].split("?")[0].split("#")[0].strip()
    return value


def is_valid_domain(domain):
    if not domain or " " in domain or "." not in domain:
        return False

    return bool(re.match(r"^[a-z0-9][a-z0-9.-]*\.[a-z]{2,}$", domain))


def normalize_email(email):
    return str(email or "").strip().lower()


def is_valid_email(email):
    value = normalize_email(email)
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]{2,}$", value))


def normalize_phone_number(phone):
    digits = re.sub(r"\D", "", str(phone or ""))
    if len(digits) == 10:
        return f"{digits[:3]} {digits[3:6]} {digits[6:8]} {digits[8:]}"
    if len(digits) == 9:
        return f"{digits[:3]} {digits[3:6]} {digits[6:]}"
    return str(phone or "").strip()


def widget_int(key):
    try:
        return int(st.session_state.get(key, 0) or 0)
    except (TypeError, ValueError):
        return 0


def get_regulators_by_industry(industry):
    regulators = {
        "Финтех / Банки": """
- Национальный Банк РК
- PCI DSS
- ISO 27001
- Постановления НБРК по ИБ
""",

        "Госсектор": """
- ГОСТ РК 34
- Требования ГТС
- Требования ИБ государственных ИС
- ISO 27001
""",

        "Ритейл / E-commerce": """
- PCI DSS
- Закон РК о персональных данных
- ISO 27001
""",

        "IT / Разработка": """
- OWASP ASVS
- Secure SDLC
- ISO 27001
- SOC2
""",

        "Производство": """
- ISA/IEC 62443
- ISO 27001
- Требования по защите АСУ ТП
"""
    }

    return regulators.get(
        industry,
        """
- ISO 27001
- Закон РК о персональных данных
"""
    )

def generate_rule_based_risks(results, context):

    risks = []
    users = context.get("users", 0)
    servers = context.get("servers", 0)

    has_critical_systems = context.get(
        "has_critical_systems",
        False
    )

    has_personal_data = context.get(
        "has_personal_data",
        False
    )

    has_public_web = context.get(
        "has_public_web",
        False
    )

    has_development = context.get(
        "has_development",
        False
    )

    large_company = context.get(
        "large_company",
        False
    )

    enterprise_company = context.get(
        "enterprise_company",
        False
    )

    users = results.get("_user_count", 0)

    # =========================
    # MFA / ACCESS CONTROL
    # =========================

    if results.get("MFA") == "Нет":
        access_scope = (
            "VPN, административных учетных записей, почты и критичных систем"
            if results.get("VPN") != "Нет"
            else "административных учетных записей, почты и критичных систем"
        )
        risks.append({
            "level": "CRITICAL",
            "risk": "Критичные доступы не защищены MFA",
            "description": f"В анкете не указана многофакторная аутентификация для {access_scope}.",
            "impact": "Компрометация одного пароля может привести к несанкционированному доступу к корпоративным системам.",
            "recommendation": "Внедрить MFA поэтапно: сначала для администраторов, VPN, почты и критичных бизнес-систем.",
            "regulators": ["ISO 27001", "NIST", "CIS Controls"],
            "vendors": ["Cisco Duo", "FortiAuthenticator", "Axidian"]
        })

    if results.get("VPN") != "Нет" and results.get("MFA") == "Нет":

        risks.append({
            "level": "CRITICAL",
            "risk": "Удаленный доступ без MFA",
            "description": "VPN-доступ реализован без многофакторной аутентификации.",
            "impact": "Высокий риск компрометации учетных записей и несанкционированного доступа.",
            "recommendation": "Внедрить MFA для VPN, административного доступа и критичных систем.",
            "regulators": ["ISO 27001", "NIST", "PCI DSS"],
            "vendors": ["Cisco Duo", "FortiAuthenticator", "Axidian"]
        })

        # =========================
    # NO SIEM
    # =========================

    if results.get("SIEM") == "Нет":

        severity = "LOW"
        recommendation = (
            "Для текущего масштаба инфраструктуры достаточно настроить "
            "централизованный сбор критичных журналов и регламент реакции. "
            "Полноценный SIEM целесообразно рассматривать при росте инфраструктуры "
            "или наличии регуляторных требований."
        )

        if (
            has_critical_systems
            or has_personal_data
            or large_company
        ):
            severity = "MEDIUM"
            recommendation = (
                "Рассмотреть MSSP/SOC или легковесный SIEM-сценарий для критичных "
                "систем, VPN, серверов и средств защиты."
            )

        if enterprise_company:
            severity = "HIGH"
            recommendation = (
                "Внедрить SIEM/SOC как целевую модель централизованного мониторинга "
                "с корреляцией событий, регламентами реагирования и регулярной "
                "аналитикой инцидентов."
            )

        risks.append({
            "level": severity,
            "risk": "Отсутствует централизованный мониторинг ИБ",
            "description": (
                "События безопасности не агрегируются "
                "в единой системе мониторинга."
            ),
            "impact": (
                "Увеличение времени обнаружения атак "
                "и расследования инцидентов."
            ),
            "recommendation": recommendation,
            "regulators": [
                "ISO 27001",
                "NIST CSF"
            ],
            "vendors": [
                "IBM QRadar",
                "Splunk",
                "R-Vision SIEM"
            ]
        })

    # =========================
    # EPP WITHOUT EDR
    # =========================

    if results.get("Антивирус") != "Нет" and results.get("EDR") == "Нет":

        risks.append({
            "level": "HIGH",
            "risk": "Endpoint-защита ограничена только антивирусом",
            "description": "Используется только базовая антивирусная защита без EDR/XDR-функциональности.",
            "impact": "Низкая эффективность обнаружения сложных атак, ransomware и lateral movement.",
            "recommendation": "Рассмотреть внедрение EDR/XDR платформы.",
            "regulators": ["NIST", "MITRE ATT&CK"],
            "vendors": ["CrowdStrike", "SentinelOne", "Trend Micro"]
        })

    if results.get("Антивирус") == "Нет" and users > 0:
        risks.append({
            "level": "CRITICAL",
            "risk": "Не указана базовая защита рабочих мест",
            "description": f"В анкете указано {users} АРМ, но не указано наличие EPP/антивирусной защиты.",
            "impact": "Рабочие станции остаются наиболее вероятной точкой входа для malware, ransomware и фишинговых атак.",
            "recommendation": "Внедрить централизованную EPP-защиту с едиными политиками, отчетностью и контролем статуса агентов.",
            "regulators": ["CIS Controls", "NIST CSF", "ISO 27001"],
            "vendors": ["Kaspersky", "Trend Micro", "Bitdefender", "Eset"]
        })

    # =========================
    # LARGE INFRA WITHOUT SEGMENTATION
    # =========================

    if users > 100 and results.get("Сегментация сети") == "Нет":

        risks.append({
            "level": "HIGH",
            "risk": "Отсутствует сегментация сети",
            "description": "Крупная инфраструктура эксплуатируется без сетевой сегментации.",
            "impact": "Высокий риск lateral movement и распространения malware между сегментами.",
            "recommendation": "Реализовать VLAN/ACL/Zero Trust сегментацию.",
            "regulators": ["ISO 27001", "NIST"],
            "vendors": ["Cisco", "Aruba", "Fortinet"]
        })

    # =========================
    # BACKUP RISKS
    # =========================

    if results.get("Резервное копирование") != "Нет" and results.get("Immutable Backup") == "Нет":

        risks.append({
            "level": "HIGH",
            "risk": "Backup не защищен от ransomware",
            "description": "Отсутствует immutable/offline backup.",
            "impact": "Риск уничтожения резервных копий при ransomware-инциденте.",
            "recommendation": "Внедрить immutable backup и air-gap копии.",
            "regulators": ["NIST", "ISO 27001"],
            "vendors": ["Veeam", "Commvault", "Rubrik"]
        })

    if results.get("Резервное копирование") == "Нет" and servers > 0:
        risks.append({
            "level": "CRITICAL",
            "risk": "Не указан контур резервного копирования серверов",
            "description": f"В инфраструктуре указано {servers} серверов, но резервное копирование не описано.",
            "impact": "При сбое, ошибке администратора или ransomware-инциденте восстановление сервисов может быть невозможным.",
            "recommendation": "Определить RPO/RTO, внедрить регулярный backup критичных систем и провести тестовое восстановление.",
            "regulators": ["ISO 27001", "NIST CSF", "CIS Controls"],
            "vendors": ["Veeam", "Commvault", "Veritas"]
        })

    if servers >= 10 and results.get("Резервное копирование") != "Нет" and results.get("DR") == "Нет":
        risks.append({
            "level": "HIGH",
            "risk": "Не описан план аварийного восстановления ИТ-сервисов",
            "description": f"В инфраструктуре указано {servers} серверов и backup-контур, но не описаны DR-сценарии, RTO/RPO и регулярные тесты восстановления.",
            "impact": "При сбое площадки, СХД или критичных виртуальных машин восстановление может занять дольше допустимого для бизнеса.",
            "recommendation": "Описать критичные сервисы, определить RTO/RPO, провести тест восстановления и подготовить поэтапный DR-план для ERP, CRM, почты и файловых сервисов.",
            "regulators": ["ISO 22301", "ISO 27001", "NIST CSF"],
            "vendors": ["Veeam", "Commvault", "Rubrik"]
        })

    if servers >= 10 and results.get("Мониторинг") == "Нет":
        risks.append({
            "level": "MEDIUM",
            "risk": "Не описан эксплуатационный мониторинг ИТ-инфраструктуры",
            "description": "В анкете есть серверы, сеть и бизнес-системы, но не указан единый мониторинг доступности, производительности и емкости.",
            "impact": "Инциденты по дискам, каналам, виртуальным машинам или бизнес-сервисам могут обнаруживаться постфактум по обращениям пользователей.",
            "recommendation": "Ввести мониторинг доступности и емкости для серверов, СХД, каналов связи, виртуализации и ключевых приложений; настроить пороги, ответственных и регулярный обзор трендов.",
            "regulators": ["ITIL", "ISO 20000", "ISO 27001"],
            "vendors": ["Zabbix", "PRTG", "ManageEngine"]
        })

    if results.get("Виртуализация") != "Нет" and results.get("Кластеризация") == "Нет":
        risks.append({
            "level": "MEDIUM",
            "risk": "Виртуализация требует проверки отказоустойчивости",
            "description": "В инфраструктуре используются виртуальные серверы, но не указаны кластеризация, правила распределения нагрузок и сценарии отказа хостов.",
            "impact": "Отказ одного физического хоста или ошибки распределения ресурсов могут повлиять сразу на несколько бизнес-сервисов.",
            "recommendation": "Проверить HA-настройки, резервы CPU/RAM/storage, правила размещения критичных VM и порядок обслуживания хостов без простоя сервисов.",
            "regulators": ["ITIL", "ISO 20000"],
            "vendors": ["VMware", "Veeam", "Zabbix"]
        })

    if results.get("СХД") != "Нет" and results.get("Мониторинг СХД") == "Нет":
        risks.append({
            "level": "MEDIUM",
            "risk": "Не описан контроль емкости и производительности СХД",
            "description": "В анкете указана СХД, но не указаны процессы контроля емкости, производительности, snapshot-политик и предупреждения деградации.",
            "impact": "Рост данных, деградация RAID-групп или нехватка производительности могут привести к снижению доступности ERP, CRM и файловых сервисов.",
            "recommendation": "Ввести регулярный capacity management для СХД: пороги заполнения, контроль latency/IOPS, проверку snapshot-политик и план расширения емкости.",
            "regulators": ["ITIL", "ISO 20000"],
            "vendors": ["Zabbix", "PRTG", "ManageEngine"]
        })

    if has_public_web and results.get("WAF") == "Нет":
        risks.append({
            "level": "HIGH",
            "risk": "Публичные веб-сервисы не защищены WAF",
            "description": "В анкете указаны публичные веб-ресурсы, но не указана WAF-защита.",
            "impact": "Повышается риск атак на веб-приложения, утечки данных и нарушения доступности сервиса.",
            "recommendation": "Провести экспресс-оценку web-периметра и внедрить WAF/CDN-защиту для критичных публичных ресурсов.",
            "regulators": ["OWASP ASVS", "PCI DSS", "ISO 27001"],
            "vendors": ["Imperva", "F5", "Cloudflare", "Radware"]
        })

    if results.get("Mail Security") == "Нет":
        risks.append({
            "level": "MEDIUM",
            "risk": "Не указана специализированная защита электронной почты",
            "description": "В анкете отсутствует информация о mail security/anti-phishing защите.",
            "impact": "Фишинг и вредоносные вложения остаются одним из основных сценариев первичного компрометации.",
            "recommendation": "Проверить текущий почтовый контур и внедрить защиту от фишинга, вредоносных вложений и подмены отправителя.",
            "regulators": ["CIS Controls", "NIST CSF"],
            "vendors": ["Trend Micro", "Forcepoint", "Barracuda"]
        })

    if results.get("PAM") == "Нет" and (servers >= 10 or has_critical_systems):
        risks.append({
            "level": "HIGH",
            "risk": "Привилегированные учетные записи не выделены в отдельный контроль",
            "description": "Есть серверы или критичные системы, но не указано использование PAM.",
            "impact": "Компрометация администраторской учетной записи может привести к полному контролю над инфраструктурой.",
            "recommendation": "Ввести учет привилегированных доступов, vault для админ-паролей и контроль сессий для критичных систем.",
            "regulators": ["ISO 27001", "CIS Controls", "NIST"],
            "vendors": ["CyberArk", "Fudo", "Netwrix"]
        })

    if has_personal_data and results.get("DLP") == "Нет":
        risks.append({
            "level": "MEDIUM",
            "risk": "Не указан контроль утечек конфиденциальных данных",
            "description": "Компания обрабатывает чувствительные данные, но DLP/контроль каналов утечки не указан.",
            "impact": "Растет риск несанкционированной передачи персональных, клиентских или коммерческих данных.",
            "recommendation": "Определить критичные типы данных и внедрить контроль основных каналов: почта, web, USB, облачные хранилища.",
            "regulators": ["Закон РК о персональных данных", "ISO 27001"],
            "vendors": ["Forcepoint", "Zecurion", "Гарда", "Symantec"]
        })

    if has_development and results.get("SAST") == "Нет" and results.get("DAST") == "Нет":
        risks.append({
            "level": "MEDIUM",
            "risk": "Безопасность приложений не встроена в разработку",
            "description": "В анкете указана разработка, но не указаны SAST/DAST-проверки.",
            "impact": "Уязвимости приложений могут попадать в продуктив и становиться точкой компрометации публичных сервисов.",
            "recommendation": "Встроить базовые SAST/DAST-проверки в CI/CD или регулярный процесс релизной проверки.",
            "regulators": ["OWASP ASVS", "ISO 27001", "PCI DSS"],
            "vendors": ["Positive Technologies", "Checkmarx", "HCL"]
        })

    # =========================
    # PATCH MANAGEMENT
    # =========================

    old_os_detected = any([
        results.get("ОС АРМ (Windows XP/Vista/7/8)", 0) > 0,
        results.get("ОС Сервера (Windows Server 2008/2012 R2)", 0) > 0
    ])

    if results.get("Patch Management") == "Нет":

        severity = "MEDIUM"

        if old_os_detected:
            severity = "CRITICAL"

        elif users > 100:
            severity = "HIGH"

        risks.append({
            "level": severity,
            "risk": "Отсутствует централизованный Patch Management",
            "description": (
                "Обновления устанавливаются "
                "несистемно либо вручную."
            ),
            "impact": (
                "Высокая вероятность эксплуатации "
                "известных уязвимостей."
            ),
            "recommendation": (
                "Внедрить централизованное "
                "управление обновлениями."
            ),
            "regulators": [
                "CIS Controls",
                "NIST"
            ],
            "vendors": [
                "ManageEngine",
                "Ivanti",
                "Tanium",
                "Microsoft MECM"
            ]
        })
    # =========================
    # LEGACY OS
    # =========================

    legacy_workstations = (
        results.get(
            "ОС АРМ (Windows XP/Vista/7/8)",
            0
        )
    )

    legacy_servers = (
        results.get(
            "ОС Сервера (Windows Server 2008/2012 R2)",
            0
        )
    )

    if legacy_workstations > 0 or legacy_servers > 0:

        risks.append({

            "level": "CRITICAL",

            "risk": "Использование устаревших операционных систем",

            "description":
                f"Обнаружены устаревшие ОС. "
                f"АРМ: {legacy_workstations}, "
                f"Серверы: {legacy_servers}.",

            "impact":
                "Уязвимости больше не исправляются "
                "производителем.",

            "recommendation":
                "Разработать программу миграции "
                "на поддерживаемые версии ОС.",

            "regulators": [
                "ISO 27001",
                "CIS Controls"
            ],

            "vendors": [
                "Microsoft",
                "Red Hat",
                "VMware",
                "Citrix"
            ]
        })

    priority_order = {"CRITICAL": 1, "HIGH": 2, "MEDIUM": 3, "LOW": 4}
    unique_risks = []
    seen = set()
    for risk in risks:
        title = risk_semantic_key(risk)
        if title and title not in seen:
            unique_risks.append(risk)
            seen.add(title)

    return sorted(
        unique_risks,
        key=lambda item: priority_order.get(str(item.get("level", "")).upper(), 99)
    )
def ai_generate_risks_and_recs(c_info, results):
    import json
    import streamlit as st

    api_key = get_app_secret("GEMINI_API_KEY")

    if not api_key:
        return []

    try:
        model_name = get_app_secret("GEMINI_MODEL", "gemini-2.5-flash")
        ai_timeout = int(get_app_secret("GEMINI_TIMEOUT_SECONDS", 45))
        fallback_models = str(get_app_secret(
            "GEMINI_FALLBACK_MODELS",
            "gemini-2.5-flash-lite"
        ))
        model_candidates = []
        for candidate in [model_name, *fallback_models.split(",")]:
            candidate = str(candidate).strip()
            if candidate and candidate not in model_candidates:
                model_candidates.append(candidate)

        safe_client, safe_results = sanitize_for_ai(
            c_info,
            results
        )
        enabled_controls, missing_controls = security_control_snapshot(results)
        enabled_controls_text = "\n".join(enabled_controls) if enabled_controls else "Не указаны"
        missing_controls_text = ", ".join(missing_controls) if missing_controls else "Нет явных пробелов"

        vendor_context = ""

        regulator_context = get_regulators_by_industry(
            c_info.get("Сфера деятельности", "")
        )

        def summarize_for_ai(values, limit=70):
            summary = []
            priority_markers = (
                "арм", "ос", "сервер", "виртуал", "схд", "резерв", "backup",
                "маршрут", "канал", "ngfw", "vpn", "epp", "edr", "xdr", "mdr",
                "dlp", "mail", "waf", "ddos", "ids", "nac", "ztna", "iam",
                "mfa", "pam", "siem", "soar", "уязв", "patch", "web", "разработ"
            )
            for key, value in values.items():
                if str(key).startswith("_"):
                    continue
                text_value = str(value).strip()
                if not text_value or text_value.lower() in {"нет", "none", "nan", "-"}:
                    continue
                key_text = str(key)
                normalized_key = key_text.lower()
                if priority_markers and not any(marker in normalized_key for marker in priority_markers):
                    continue
                summary.append(f"{key_text}: {text_value[:180]}")
                if len(summary) >= limit:
                    break
            if not summary:
                for key, value in values.items():
                    if str(key).startswith("_"):
                        continue
                    text_value = str(value).strip()
                    if text_value:
                        summary.append(f"{key}: {text_value[:180]}")
                    if len(summary) >= 35:
                        break
            return "\n".join(summary)

        ai_summary = summarize_for_ai(safe_results)

        prompt = f"""
Выступай как аудитор Big4, CISO, CTO, Enterprise Security Architect и руководитель ИТ-инфраструктуры.

ЗАДАЧА:
Провести профессиональный анализ ИТ и ИБ инфраструктуры компании на основании опросника.
Отчет должен включать не только кибербезопасность, но и ИТ-управляемость: сеть, серверы, виртуализацию, СХД, backup, DR, мониторинг, эксплуатационные процессы, бизнес-системы и разработку.

ВАЖНЫЕ ПРАВИЛА:

1. Не придумывай технологии без оснований в данных.

2. Не создавай риски только потому, что отсутствует продукт.

3. Оценивай бизнес-контекст компании.

4. Для каждой рекомендации сначала определяй КАТЕГОРИЮ решения:

Примеры:

Отсутствует SIEM
→ категория SIEM

Отсутствует Patch Management
→ категория Patch Management

Отсутствует Backup
→ категория Backup

Отсутствует CASB
→ категория CASB

Отсутствует ZTNA
→ категория ZTNA

Отсутствует SAST
→ категория Application Security

Устаревшие ОС
→ категория Migration Project

Устаревшие серверы
→ категория Infrastructure Modernization

5. Для категории решения сначала используй вендоров из каталога.

6. Если в каталоге отсутствуют подходящие решения,
разрешается использовать мировых лидеров рынка.

7. Никогда не предлагай:
- антивирусы
- EDR
- NGFW
- DLP

для устранения проблем устаревших операционных систем.

8. Для Legacy OS основная рекомендация:

- миграция
- модернизация
- виртуализация
- сегментация
- изоляция

9. Не предлагай SOAR как критическую необходимость
для компаний без зрелого SOC.

10. Не предлагай ZTNA как высокий риск без признаков удаленного доступа.

11. Для каждого риска указывай не более 3 наиболее релевантных решений.

12. Используй вендоров из каталога как приоритетных.

13. Если подходящего решения в каталоге нет,
подбери наиболее распространенные мировые решения.

14. Не дублируй риски.

15. Верни 10-12 наиболее важных рисков, если данных достаточно. Если реальных рисков меньше,
не растягивай отчет искусственно.

16. Приоритет риска определяй по шкале:

CRITICAL
HIGH
MEDIUM
LOW

17. Особое внимание уделяй:

- устаревшим ОС
- отсутствию резервного копирования
- отсутствию MFA
- отсутствию Patch Management
- отсутствию сегментации сети
- отсутствию мониторинга безопасности
- отсутствию защиты почты
- отсутствию контроля привилегированных учетных записей

18. Цель отчета - экспертная диагностика для клиента. Формулировки должны звучать как
профессиональное аудиторское заключение, а не как коммерческое предложение.

19. Рекомендации должны учитывать масштаб инфраструктуры:
- для малой инфраструктуры не предлагай тяжелый SIEM/SOAR как первоочередной проект;
- для средней инфраструктуры допускай MSSP/SOC, управляемые сервисы и поэтапные внедрения;
- для крупной инфраструктуры допускай полноценные платформенные проекты.

20. Не перечисляй все отсутствующие технологии как риски. Выбирай только то, что логично
для размера, серверов, публичных сервисов, бизнес-систем, разработки и текущих средств защиты.

21. Для каждого риска дай конкретную, выполнимую рекомендацию и объясни, какой риск
она снижает.

22. Если в данных есть сеть, серверы, СХД, виртуализация, backup, бизнес-системы или разработка,
обязательно включи 3-5 ИТ-рекомендаций, а не только ИБ-рекомендации.

23. Не используй Microsoft как ИБ-вендора. Microsoft допустим только для тем ОС,
миграции Windows, Windows Server или обновления устаревших систем.

24. Не создавай несколько отдельных пунктов про EDR/XDR/MDR. Если тема одна,
объедини ее в один зрелый риск по защите рабочих мест и реагированию.

25. Не рекомендуй внедрение контроля, если он уже указан в блоке "УЖЕ ВНЕДРЕНО".
Если контроль есть, можно рекомендовать только улучшение покрытия, регламент,
контроль исключений или проверку эффективности, но нельзя писать, что контроль отсутствует.

УЖЕ ВНЕДРЕНО:
{enabled_controls_text}

НЕ УКАЗАНО В АНКЕТЕ:
{missing_controls_text}

ДАННЫЕ АУДИТА:

{safe_results}

СФЕРА ДЕЯТЕЛЬНОСТИ:

{c_info.get("Сфера деятельности", "")}

КАТАЛОГ РЕШЕНИЙ КОМПАНИИ:

{vendor_context}

РЕГУЛЯТОРНЫЕ ТРЕБОВАНИЯ:

{regulator_context}

Верни ТОЛЬКО JSON.

Формат:

[
  {{
    "level": "HIGH",
    "risk": "Название риска",
    "description": "Описание риска",
    "impact": "Последствия для бизнеса",
    "recommendation": "Что необходимо сделать",
    "vendors": [
      "Vendor1",
      "Vendor2",
      "Vendor3"
    ],
    "regulators": [
      "ISO 27001"
    ]
  }}
]
"""

        def gemini_url(active_model):
            return (
                "https://generativelanguage.googleapis.com/v1beta/models/"
                f"{active_model}:generateContent"
            )

        compact_prompt = f"""
Ты CISO/CTO-аудитор. По обезличенной анкете верни только JSON-объект для экспертного отчета ИТ и ИБ.
Без markdown, без пояснений вне JSON. Первый символ ответа должен быть {{.
JSON должен быть валидным: все строковые значения в одну строку, без переносов внутри кавычек.
Если нужно разделить мысль, используй точку с пробелом, а не перевод строки.

Строгий формат ответа:
{{
  "executive_summary": [
    "5-7 содержательных выводов для руководства без коммерческого тона"
  ],
  "audit_observations": [
    {{"title": "Короткий домен", "text": "Наблюдение аудитора, связанное с фактами анкеты"}}
  ],
  "it_recommendations": [
    {{
      "level": "HIGH",
      "risk": "Риск или зона незрелости ИТ",
      "description": "1-2 предложения с фактами анкеты",
      "impact": "Бизнес-последствие",
      "recommendation": "3 конкретных шага: быстрый шаг, проектный шаг, контроль результата",
      "vendors": ["категория или решение"],
      "regulators": ["ITIL"]
    }}
  ],
  "security_recommendations": [
    {{
      "level": "HIGH",
      "risk": "Риск или зона незрелости ИБ",
      "description": "1-2 предложения с фактами анкеты",
      "impact": "Бизнес-последствие",
      "recommendation": "3 конкретных шага: быстрый шаг, проектный шаг, контроль результата",
      "vendors": ["категория или решение"],
      "regulators": ["ISO 27001"]
    }}
  ],
  "management_decisions": [
    "4-6 решений, которые руководитель может утвердить сразу"
  ],
  "roadmap": [
    {{"phase": "0-30 дней", "priority": "P1", "domain": "ИТ/ИБ", "action": "Что сделать", "rationale": "Почему это важно", "owner": "ИТ/ИБ", "effort": "Низкая|Средняя|Высокая"}}
  ]
}}

Правила:
- учитывай масштаб инфраструктуры;
- верни 5-6 it_recommendations и 5-7 security_recommendations;
- ИТ-рекомендации должны быть про сеть, каналы, серверы, виртуализацию, СХД, backup/DR, мониторинг, patch/change/capacity management, бизнес-системы и разработку;
- минимум 4 ИТ-рекомендации должны быть самостоятельными и не сводиться к ИБ-продуктам;
- ИБ-рекомендации должны учитывать уже внедренные контроли и не повторять отсутствующий контроль, если он есть;
- каждое наблюдение должно связывать минимум два факта анкеты, если это возможно;
- избегай примитивных фраз уровня "внедрить продукт"; сначала опиши управленческую/техническую меру, затем только категорию решения;
- не повторяй один и тот же домен разными словами: MFA, удаленный доступ и учетные записи объединяй в один сильный риск, если нет отдельного факта;
- не повторяй EDR/XDR/MDR несколькими пунктами; объединяй в один риск по endpoint detection/response;
- для малого масштаба не предлагай тяжелый SIEM/SOAR как первоочередной проект;
- если инфраструктура малая, сначала предлагай базовые процессы, MFA, backup, patch management, EPP/EDR по необходимости и сегментацию;
- если инфраструктура средняя или крупная, оцени связки: учетные записи, сеть, серверы, резервное копирование, почта, уязвимости, мониторинг и реагирование;
- делай рекомендации экспертными: что сделать, зачем, какой ожидаемый эффект;
- для каждой рекомендации дай порядок внедрения: быстрый шаг, проектный шаг, контроль результата;
- legacy OS закрывай миграцией, изоляцией, сегментацией, обновлением, а не EDR/DLP;
- не перечисляй все отсутствующие продукты как риски;
- не путай категории: MFA не закрывается PAM, уязвимости не закрываются SIEM, устаревшие ОС не закрываются EDR.
- не используй Microsoft как ИБ-вендора; Microsoft допустим только для ОС/миграции Windows/Windows Server.
- строго не пиши "отсутствует", "не внедрено" или "не указано" про контроль из списка "Уже внедрено";
- если MFA есть в списке "Уже внедрено", не создавай риск по отсутствию MFA.

Отрасль: {c_info.get("Сфера деятельности", "-")}

Уже внедрено:
{enabled_controls_text}

Не указано:
{missing_controls_text}

Ключевые данные анкеты:
{ai_summary}

Регуляторный контекст:
{regulator_context[:1200]}
"""
        fallback_payload = {
            "contents": [
                {
                    "role": "user",
                    "parts": [{"text": compact_prompt}]
                }
            ],
            "generationConfig": {
                "temperature": 0.15,
                "maxOutputTokens": 6144,
            },
        }

        minimal_payload = {
            "contents": [
                {
                    "role": "user",
                    "parts": [{"text": compact_prompt[:6000]}]
                }
            ],
            "generationConfig": {
                "temperature": 0.05,
                "maxOutputTokens": 4096,
            },
        }

        line_prompt = f"""
Ты CISO/CTO-аудитор. Верни 8-10 строк без JSON и без markdown. Покрой ИТ и ИБ, не только endpoint.
Формат каждой строки строго:
LEVEL | RISK | DESCRIPTION | IMPACT | RECOMMENDATION

LEVEL только CRITICAL, HIGH, MEDIUM или LOW.
Не используй символ | внутри полей.
Каждое поле в одну строку.
Не используй Microsoft как ИБ-вендора. Не дублируй EDR/XDR/MDR отдельными строками.
Не пиши, что отсутствует контроль из списка "Уже внедрено".

Уже внедрено:
{enabled_controls_text}

Не указано:
{missing_controls_text}

Данные:
{ai_summary}
"""
        line_payload = {
            "contents": [
                {
                    "role": "user",
                    "parts": [{"text": line_prompt[:5000]}]
                }
            ],
            "generationConfig": {
                "temperature": 0.05,
                "maxOutputTokens": 4096,
            },
        }

        def focused_line_payload(title, target_count, focus):
            focused_prompt = f"""
Ты CISO/CTO-аудитор. Верни ровно {target_count} строк без JSON и без markdown.
Формат каждой строки строго:
LEVEL | RISK | DESCRIPTION | IMPACT | RECOMMENDATION

LEVEL только CRITICAL, HIGH, MEDIUM или LOW.
Не используй символ | внутри полей.
Каждое поле в одну строку, без нумерации.
Каждая рекомендация должна быть законченной фразой: быстрый шаг; проектный шаг; контроль результата.
Не используй Microsoft как ИБ-вендора. Microsoft допустим только для ОС/Windows/Windows Server.
Не дублируй EDR/XDR/MDR отдельными строками.

Раздел: {title}
Фокус анализа: {focus}

Отрасль: {c_info.get("Сфера деятельности", "-")}

Ключевые данные анкеты:
{ai_summary}
"""
            return {
                "contents": [
                    {
                        "role": "user",
                        "parts": [{"text": focused_prompt[:5200]}]
                    }
                ],
                "generationConfig": {
                    "temperature": 0.05,
                    "maxOutputTokens": 3072,
                },
            }

        focused_payloads = (
            (
                "line",
                focused_line_payload(
                    "ИТ-инфраструктура и эксплуатационная зрелость",
                    4,
                    "сеть, каналы связи, серверы, виртуализация, СХД, backup, DR, мониторинг, бизнес-системы, разработка"
                )
            ),
            (
                "line",
                focused_line_payload(
                    "Информационная безопасность",
                    5,
                    "MFA, endpoint detection/response, vulnerability/patch management, mail security, WAF, PAM, SIEM/SOC, DLP"
                )
            ),
        )

        def extract_gemini_error(response):
            try:
                payload = response.json()
                message = payload.get("error", {}).get("message", response.text)
                status = payload.get("error", {}).get("status", "")
                return f"HTTP {response.status_code} {status}: {message}"
            except Exception:
                return f"HTTP {response.status_code}: {response.text[:1200]}"

        def call_gemini(request_payload, active_model):
            active_url = gemini_url(active_model)
            response_payload = None
            if os.name != "nt":
                import requests

                def gemini_post(verify):
                    return requests.post(
                        active_url,
                        params={"key": api_key},
                        json=request_payload,
                        timeout=ai_timeout,
                        verify=verify,
                    )

                try:
                    response = gemini_post(REQUEST_VERIFY)
                except requests.exceptions.SSLError:
                    response = gemini_post(False)
                if not response.ok:
                    raise RuntimeError(extract_gemini_error(response))
                return response.json()

            if response_payload is None:
                response_payload = node_fetch_json(
                    f"{active_url}?key={api_key}",
                    {
                        "url": f"{active_url}?key={api_key}",
                        "method": "POST",
                        "body": request_payload,
                        "timeoutMs": ai_timeout * 1000,
                    },
                    timeout_seconds=ai_timeout,
                )
            return response_payload

        def gemini_response_text(response_payload):
            parts = (
                response_payload
                .get("candidates", [{}])[0]
                .get("content", {})
                .get("parts", [])
            )
            return "\n".join(
                str(part.get("text", ""))
                for part in parts
                if part.get("text")
            ).strip()

        def gemini_empty_reason(response_payload):
            candidate = response_payload.get("candidates", [{}])[0]
            finish_reason = candidate.get("finishReason", "UNKNOWN")
            prompt_feedback = response_payload.get("promptFeedback", {})
            return f"пустой ответ Gemini; finishReason={finish_reason}; promptFeedback={prompt_feedback}"

        gemini_retry_count = int(get_app_secret("GEMINI_RETRY_COUNT", 1))
        gemini_retry_delay = float(get_app_secret("GEMINI_RETRY_DELAY_SECONDS", 2.0))

        def should_retry_gemini_error(error_text):
            lowered = str(error_text or "").lower()
            return any(
                marker in lowered
                for marker in (
                    "503",
                    "unavailable",
                    "high demand",
                    "пустой ответ gemini",
                    "finishreason=unknown",
                )
            )

        def call_gemini_with_retries(request_payload, active_model):
            last_error = None
            for attempt in range(gemini_retry_count + 1):
                try:
                    response_payload = call_gemini(request_payload, active_model)
                    response_text = gemini_response_text(response_payload)
                    if response_text:
                        return response_payload, response_text

                    last_error = gemini_empty_reason(response_payload)
                except Exception as exc:
                    last_error = redact_secret(exc, api_key)

                if attempt < gemini_retry_count and should_retry_gemini_error(last_error):
                    time.sleep(gemini_retry_delay * (attempt + 1))
                    continue

                break

            raise RuntimeError(str(last_error or "Gemini did not return text"))

        def repair_json_text(text):
            cleaned = []
            in_string = False
            escaped = False
            for char in text:
                if in_string and char in "\r\n\t":
                    cleaned.append(" ")
                    escaped = False
                    continue
                cleaned.append(char)
                if escaped:
                    escaped = False
                elif char == "\\":
                    escaped = True
                elif char == '"':
                    in_string = not in_string

            repaired = "".join(cleaned).strip()
            array_start = repaired.find("[")
            object_start = repaired.find("{")
            if object_start >= 0 and (array_start < 0 or object_start < array_start):
                start = object_start
                end = repaired.rfind("}")
            else:
                start = array_start
                end = repaired.rfind("]")
            if start >= 0:
                repaired = repaired[start:end + 1] if end > start else repaired[start:]

            open_strings = repaired.count('"') % 2
            if open_strings:
                repaired += '"'

            open_braces = repaired.count("{") - repaired.count("}")
            open_brackets = repaired.count("[") - repaired.count("]")
            if open_braces > 0:
                repaired += "}" * open_braces
            if open_brackets > 0:
                repaired += "]" * open_brackets

            repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
            return repaired

        def parse_ai_response_text(response_text):
            response_text = re.sub(r"^```(?:json)?", "", response_text).strip()
            response_text = re.sub(r"```$", "", response_text).strip()
            variants = [response_text]
            start = response_text.find("[")
            end = response_text.rfind("]")
            if start >= 0 and end > start:
                variants.append(response_text[start:end + 1])
            object_start = response_text.find("{")
            object_end = response_text.rfind("}")
            if object_start >= 0 and object_end > object_start:
                variants.append(response_text[object_start:object_end + 1])
            variants.append(repair_json_text(response_text))

            last_error = None
            for variant in variants:
                try:
                    return json.loads(variant)
                except json.JSONDecodeError as exc:
                    last_error = exc
            raise last_error

        def parse_line_response(response_text):
            rows = []
            for raw_line in response_text.splitlines():
                line = raw_line.strip().lstrip("-•0123456789. ")
                if "|" not in line:
                    continue
                parts = [part.strip() for part in line.split("|")]
                if len(parts) < 5:
                    continue
                level, risk, description, impact, recommendation = parts[:5]
                if level.upper() not in {"CRITICAL", "HIGH", "MEDIUM", "LOW"}:
                    continue
                rows.append({
                    "level": level.upper(),
                    "risk": risk,
                    "description": description,
                    "impact": impact,
                    "recommendation": recommendation,
                    "vendors": [],
                    "regulators": [],
                })
            return rows

        ai_errors = []
        payload_attempts = (
            ("json", fallback_payload),
            ("json", minimal_payload),
            ("line", line_payload),
        )

        for response_format, request_payload in payload_attempts:
            for active_model in model_candidates:
                try:
                    response_payload, response_text = call_gemini_with_retries(
                        request_payload,
                        active_model
                    )

                    try:
                        parsed_payload = (
                            parse_line_response(response_text)
                            if response_format == "line"
                            else parse_ai_response_text(response_text)
                        )
                    except json.JSONDecodeError as exc:
                        ai_errors.append(f"{active_model}: JSON parse error: {exc}")
                        continue

                    ai_narrative = normalize_ai_audit_narrative(parsed_payload)
                    normalized_payload = normalize_ai_risks_payload(parsed_payload)
                    normalized_payload, skipped_ai_items = filter_ai_risks_by_answers(
                        normalized_payload,
                        results
                    )
                    if skipped_ai_items:
                        ai_errors.append(
                            f"{active_model}: отброшены противоречивые пункты: "
                            + "; ".join(skipped_ai_items[:5])
                        )
                    prepared_payload, quality_error = ai_quality_gate(normalized_payload)
                    if prepared_payload:
                        st.session_state.ai_last_error = ""
                        st.session_state.ai_model_used = active_model
                        st.session_state.ai_audit_narrative = ai_narrative
                        return prepared_payload

                    ai_errors.append(f"{active_model}: {quality_error or 'нет пригодных законченных рекомендаций'}")
                except Exception as exc:
                    ai_errors.append(f"{active_model}: {redact_secret(exc, api_key)}")

        raise ValueError("Gemini не дал пригодный ответ. " + " | ".join(ai_errors[-6:]))

    except Exception as e:
        st.session_state.ai_last_error = redact_secret(e, api_key)
        st.session_state.ai_audit_narrative = {}
        return []



# --- AI BLOCK END ---

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

def inject_audit_design():
    st.markdown("""
    <style>
    :root {
        --audit-bg: #f6f7f9;
        --audit-panel: #ffffff;
        --audit-border: #d8dee8;
        --audit-text: #151922;
        --audit-muted: #667085;
        --audit-accent: #0f766e;
        --audit-accent-soft: #dff3ef;
        --audit-warn: #b45309;
        --audit-warn-soft: #fff7ed;
        --audit-risk: #b42318;
        --audit-risk-soft: #fff1f0;
        --audit-ink: #1f2937;
    }

    .stApp {
        background: var(--audit-bg);
        color: var(--audit-text);
    }

    section[data-testid="stSidebar"] {
        background: #fbfcfe;
        border-right: 1px solid var(--audit-border);
    }

    div[data-testid="stHeader"] {
        background: rgba(246, 247, 249, 0.92);
        backdrop-filter: blur(10px);
    }

    .block-container {
        padding-top: 2.2rem;
        padding-bottom: 7rem;
        max-width: 1380px;
    }

    h1, h2, h3 {
        letter-spacing: 0;
    }

    .audit-hero {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 22px 24px;
        margin: 0 0 14px 0;
        box-shadow: 0 10px 30px rgba(16, 24, 40, 0.05);
    }

    .audit-brand-hero {
        display: grid;
        grid-template-columns: minmax(0, 1fr) minmax(300px, 380px);
        gap: 26px;
        align-items: center;
    }

    .audit-kicker {
        color: var(--audit-accent);
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0;
        margin-bottom: 8px;
    }

    .audit-title {
        color: var(--audit-text);
        font-size: 34px;
        font-weight: 760;
        line-height: 1.12;
        margin: 0 0 8px 0;
    }

    .audit-subtitle {
        color: var(--audit-muted);
        font-size: 15px;
        line-height: 1.55;
        max-width: 900px;
        margin: 0;
    }

    .audit-start-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 10px;
        margin-top: 18px;
        max-width: 920px;
    }

    .audit-start-card {
        background: #f8fafc;
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 12px 13px;
        min-height: 86px;
    }

    .audit-start-card strong {
        display: block;
        color: var(--audit-text);
        font-size: 13px;
        margin-bottom: 5px;
    }

    .audit-start-card span {
        color: var(--audit-muted);
        display: block;
        font-size: 12px;
        line-height: 1.45;
    }

    .audit-logo-lockup {
        border-left: 1px solid var(--audit-border);
        padding-left: 24px;
        text-align: center;
    }

    .audit-logo-lockup img {
        max-width: 320px;
        max-height: 120px;
        width: auto;
        height: auto;
        object-fit: contain;
    }

    .audit-logo-lockup .brand-name {
        color: var(--audit-text);
        font-size: 13px;
        font-weight: 760;
        margin-top: 10px;
    }

    .audit-logo-lockup .brand-signature {
        color: var(--audit-accent);
        font-size: 12px;
        font-weight: 760;
        margin-top: 3px;
    }

    .audit-section {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 16px 18px;
        margin: 24px 0 14px 0;
    }

    .audit-anchor {
        display: block;
        position: relative;
        top: -84px;
        visibility: hidden;
    }

    .audit-section .eyebrow {
        color: var(--audit-muted);
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 4px;
    }

    .audit-section .title {
        color: var(--audit-text);
        font-size: 22px;
        font-weight: 760;
        margin-bottom: 4px;
    }

    .audit-section .body {
        color: var(--audit-muted);
        font-size: 14px;
        margin: 0;
    }

    .metric-card {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 14px 14px 12px 14px;
        min-height: 96px;
    }

    .metric-card .label {
        color: var(--audit-muted);
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 8px;
    }

    .metric-card .value {
        color: var(--audit-text);
        font-size: 26px;
        font-weight: 760;
        line-height: 1.05;
    }

    .metric-card .hint {
        color: var(--audit-muted);
        font-size: 12px;
        margin-top: 6px;
    }

    .domain-row {
        display: grid;
        grid-template-columns: repeat(5, minmax(0, 1fr));
        gap: 10px;
        margin: 12px 0 18px 0;
    }

    .domain-card {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 12px;
        min-height: 86px;
    }

    .domain-card strong {
        display: block;
        color: var(--audit-text);
        font-size: 13px;
        margin-bottom: 10px;
    }

    .domain-score {
        color: var(--audit-accent);
        font-size: 24px;
        font-weight: 760;
    }

    .risk-chip {
        border-radius: 8px;
        padding: 10px 12px;
        border: 1px solid var(--audit-border);
        background: var(--audit-panel);
        color: var(--audit-ink);
        margin-bottom: 8px;
        font-size: 13px;
        line-height: 1.45;
    }

    .risk-chip.critical {
        background: var(--audit-risk-soft);
        border-color: #fecdca;
    }

    .risk-chip.warn {
        background: var(--audit-warn-soft);
        border-color: #fed7aa;
    }

    .analysis-teaser {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 16px 18px;
        margin: 20px 0 12px 0;
        box-shadow: 0 10px 26px rgba(16, 24, 40, 0.04);
    }

    .analysis-teaser-head {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 18px;
        margin-bottom: 14px;
    }

    .analysis-teaser-title {
        color: var(--audit-text);
        font-size: 19px;
        font-weight: 760;
        line-height: 1.25;
        margin-bottom: 4px;
    }

    .analysis-teaser-copy {
        color: var(--audit-muted);
        font-size: 13px;
        line-height: 1.45;
        max-width: 760px;
    }

    .analysis-pill {
        background: var(--audit-accent-soft);
        color: #0f5f59;
        border: 1px solid #b9e5de;
        border-radius: 999px;
        padding: 7px 10px;
        font-size: 12px;
        font-weight: 760;
        white-space: nowrap;
    }

    .teaser-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 10px;
        margin-bottom: 14px;
    }

    .teaser-card {
        background: #f8fafc;
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 12px;
        min-height: 86px;
    }

    .teaser-card .label {
        color: var(--audit-muted);
        font-size: 12px;
        font-weight: 700;
        margin-bottom: 7px;
    }

    .teaser-card .value {
        color: var(--audit-text);
        font-size: 23px;
        font-weight: 760;
        line-height: 1.05;
    }

    .teaser-card .hint {
        color: var(--audit-muted);
        font-size: 12px;
        margin-top: 6px;
    }

    .teaser-columns {
        display: grid;
        grid-template-columns: 1.1fr 0.9fr;
        gap: 12px;
    }

    .teaser-columns .label {
        color: var(--audit-muted);
        font-size: 12px;
        font-weight: 700;
        margin-bottom: 8px;
    }

    .mini-signal {
        background: #ffffff;
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 10px 12px;
        margin-bottom: 8px;
        color: var(--audit-ink);
        font-size: 13px;
        line-height: 1.45;
    }

    .mini-signal.critical {
        border-color: #fecdca;
        background: var(--audit-risk-soft);
    }

    .mini-signal.warn {
        border-color: #fed7aa;
        background: var(--audit-warn-soft);
    }

    .mini-domain {
        display: flex;
        justify-content: space-between;
        gap: 10px;
        border-bottom: 1px solid #eef2f6;
        padding: 9px 0;
        color: var(--audit-ink);
        font-size: 13px;
    }

    .mini-domain:last-child {
        border-bottom: 0;
    }

    .mini-domain strong {
        color: var(--audit-text);
    }

    .analysis-teaser-note {
        color: var(--audit-muted);
        font-size: 12px;
        margin-top: 10px;
    }

    .analysis-status-panel {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-left: 4px solid var(--audit-accent);
        border-radius: 8px;
        padding: 16px 18px;
        margin-bottom: 16px;
        color: var(--audit-ink);
    }

    .analysis-status-title {
        color: var(--audit-text);
        font-size: 17px;
        font-weight: 760;
        margin-bottom: 6px;
    }

    .page-lock-note {
        color: var(--audit-risk);
        font-size: 13px;
        font-weight: 700;
        margin-top: 8px;
    }

    .analysis-log {
        background: #f8fafc;
        color: #344054;
        font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
        padding: 13px 14px;
        border: 1px solid var(--audit-border);
        min-height: 110px;
        border-radius: 8px;
        margin-bottom: 16px;
        font-size: 13px;
        line-height: 1.55;
    }

    .facts-panel {
        background: #fff8ed;
        border: 1px solid #fed7aa;
        border-left: 4px solid var(--audit-warn);
        border-radius: 8px;
        color: #1d2939;
        padding: 16px 18px;
        margin: 12px 0 18px 0;
        line-height: 1.55;
    }

    .facts-panel strong {
        color: #111827;
    }

    .section-feedback {
        background: #fff7ed;
        border: 1px solid #fed7aa;
        border-left: 4px solid var(--audit-warn);
        border-radius: 8px;
        color: #1d2939;
        margin: 12px 0 16px 0;
        padding: 12px 14px;
    }

    .section-feedback.ok {
        background: #f0fdf9;
        border-color: #99f6e4;
        border-left-color: var(--audit-accent);
    }

    .section-feedback-title {
        color: var(--audit-text);
        font-size: 13px;
        font-weight: 760;
        margin-bottom: 6px;
    }

    .section-feedback ul {
        margin: 0;
        padding-left: 18px;
    }

    .section-feedback li {
        font-size: 13px;
        line-height: 1.45;
        margin: 3px 0;
    }

    .sidebar-step {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 7px 0;
        color: #344054;
        font-size: 13px;
    }

    .sidebar-step-link {
        color: #344054 !important;
        text-decoration: none !important;
        border-radius: 6px;
        margin: 0 -6px;
        padding: 0 6px;
    }

    .sidebar-step-link:hover {
        background: #f2f4f7;
        color: var(--audit-text) !important;
    }

    .sidebar-dot {
        width: 9px;
        height: 9px;
        border-radius: 50%;
        flex: 0 0 9px;
    }

    .sidebar-dot.green {
        background: var(--audit-accent);
    }

    .sidebar-dot.red {
        background: var(--audit-risk);
    }

    .sidebar-dot.gray {
        background: #98a2b3;
    }

    .sidebar-error-link {
        display: block;
        color: #7a2e0e !important;
        text-decoration: none !important;
        background: #fff7ed;
        border: 1px solid #fed7aa;
        border-radius: 8px;
        padding: 8px 9px;
        margin: 6px 0;
        font-size: 12px;
        line-height: 1.35;
    }

    .sidebar-error-link:hover {
        border-color: var(--audit-warn);
        background: #ffedd5;
    }

    .sidebar-error-link span {
        display: block;
        color: #9a3412;
        font-weight: 760;
        margin-bottom: 2px;
    }

    .validation-panel {
        background: #fff7ed;
        border: 1px solid #fed7aa;
        border-left: 4px solid var(--audit-warn);
        border-radius: 8px;
        padding: 14px 16px;
        margin: 12px 0 18px 0;
    }

    .validation-panel-title {
        color: var(--audit-text);
        font-size: 16px;
        font-weight: 760;
        margin-bottom: 4px;
    }

    .validation-panel-copy {
        color: #7a2e0e;
        font-size: 13px;
        line-height: 1.45;
        margin-bottom: 12px;
    }

    .validation-error-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 8px;
    }

    .validation-error-link {
        color: #7a2e0e !important;
        text-decoration: none !important;
        background: #ffffff;
        border: 1px solid #fed7aa;
        border-radius: 8px;
        padding: 10px 12px;
        font-size: 13px;
        line-height: 1.4;
    }

    .validation-error-link:hover {
        border-color: var(--audit-warn);
        background: #fffbeb;
    }

    .validation-error-link strong {
        display: block;
        color: #9a3412;
        font-size: 12px;
        margin-bottom: 3px;
    }


    .st-key-floating_draft_save {
        position: fixed;
        left: 50%;
        bottom: 18px;
        transform: translateX(-50%);
        z-index: 9999;
        width: min(560px, calc(100vw - 32px));
        padding: 0;
        background: transparent;
        border: 0;
        box-shadow: none;
    }

    .st-key-floating_draft_save .element-container,
    .st-key-floating_draft_save iframe {
        display: block !important;
        width: 100% !important;
        min-width: 0 !important;
        height: 96px !important;
        min-height: 96px !important;
        margin: 0 !important;
    }

    @media (max-width: 900px) {
        .audit-title {
            font-size: 26px;
        }

        .audit-brand-hero {
            grid-template-columns: 1fr;
        }

        .audit-logo-lockup {
            border-left: 0;
            border-top: 1px solid var(--audit-border);
            padding-left: 0;
            padding-top: 14px;
            text-align: center;
        }

        .audit-logo-lockup img {
            max-width: min(100%, 330px);
            max-height: 126px;
        }

        .audit-start-grid,
        .teaser-grid,
        .teaser-columns {
            grid-template-columns: 1fr;
        }

        .analysis-teaser-head {
            display: block;
        }

        .analysis-pill {
            display: inline-block;
            margin-top: 10px;
        }

        .domain-row {
            grid-template-columns: 1fr;
        }

        .validation-error-grid {
            grid-template-columns: 1fr;
        }

        .st-key-floating_draft_save {
            bottom: 12px;
            width: calc(100vw - 24px);
        }
    }
    </style>
    """, unsafe_allow_html=True)


def get_logo_data_uri(path="logo.png"):
    if not os.path.exists(path):
        return ""

    with open(path, "rb") as logo_file:
        encoded = base64.b64encode(logo_file.read()).decode("ascii")

    return f"data:image/png;base64,{encoded}"


def render_app_header():
    logo_uri = get_logo_data_uri()
    logo_html = (
        f'<img src="{logo_uri}" alt="Khalil Trade">'
        if logo_uri
        else '<strong>Khalil Trade</strong>'
    )

    st.markdown(f"""
    <div class="audit-hero audit-brand-hero">
        <div>
            <div class="audit-kicker">Технический ИТ- и ИБ-аудит</div>
            <div class="audit-title">Опросник технического аудита ИТ и ИБ</div>
            <p class="audit-subtitle">
                Структурированный сбор данных для экспертной оценки инфраструктуры,
                зрелости защиты и подготовки XLSX-отчета.
            </p>
            <div class="audit-start-grid">
                <div class="audit-start-card">
                    <strong>7-10 минут на заполнение</strong>
                    <span>Начните с компании и конечных точек, остальные блоки включайте по факту наличия.</span>
                </div>
                <div class="audit-start-card">
                    <strong>Живая оценка по ходу анкеты</strong>
                    <span>Навигатор показывает готовность, слабые места и разделы, которые требуют внимания.</span>
                </div>
                <div class="audit-start-card">
                    <strong>XLSX-отчет для обсуждения</strong>
                    <span>На выходе формируется экспертная сводка, риски, домены защиты и быстрые улучшения.</span>
                </div>
            </div>
        </div>
        <div class="audit-logo-lockup">
            {logo_html}
            <div class="brand-name">Khalil Audit System v11.01</div>
            <div class="brand-signature">by Ivan Rudoy</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_anchor(anchor_id):
    if anchor_id:
        st.markdown(
            f'<span id="{html.escape(anchor_id)}" class="audit-anchor"></span>',
            unsafe_allow_html=True
        )


def render_section_marker(kicker, title, body, anchor_id=None):
    render_anchor(anchor_id)
    st.markdown(f"""
    <div class="audit-section">
        <div class="eyebrow">{html.escape(kicker)}</div>
        <div class="title">{html.escape(title)}</div>
        <p class="body">{html.escape(body)}</p>
    </div>
    """, unsafe_allow_html=True)


def get_section_errors(validation_errors, *markers):
    selected = []
    marker_values = [str(marker).lower() for marker in markers]

    for error in validation_errors:
        error_text = str(error)
        normalized = error_text.lower()
        if any(marker in normalized for marker in marker_values):
            selected.append(error_text)

    return list(dict.fromkeys(selected))


def get_error_anchor(error):
    normalized = str(error).lower()
    marker_map = [
        (("общая информация", "обязательные поля"), "section-company"),
        (("сайт", "email", "телефон"), "section-company"),
        (("арм", "ос на арм"), "section-endpoints"),
        (("маршрутизац", "маршрутизатор", "коммутатор", "core", "распредел", "wi-fi", "ngfw"), "section-network"),
        (("сервер", "хост", "резервного копирования"), "section-servers"),
        (("схд", "raid"), "section-storage"),
        (("название и версию", "версию"), "section-systems"),
        (("производитель", "иб"), "section-security"),
        (("разработчик", "языки разработки"), "section-dev"),
    ]

    for markers, anchor_id in marker_map:
        if any(marker in normalized for marker in markers):
            return anchor_id

    return "section-company"


def render_validation_summary(validation_errors):
    unique_errors = list(dict.fromkeys(validation_errors))
    if not unique_errors:
        return

    items = "".join(
        f'<a class="validation-error-link" href="#{html.escape(get_error_anchor(error))}">'
        f'<strong>Исправить</strong>{html.escape(error)}'
        f'</a>'
        for error in unique_errors
    )

    st.markdown(f"""
    <div class="validation-panel">
        <div class="validation-panel-title">Отчет пока нельзя сформировать</div>
        <div class="validation-panel-copy">
            Осталось исправить {len(unique_errors)} пункт(ов). Нажмите на карточку, чтобы перейти к нужному разделу анкеты.
        </div>
        <div class="validation-error-grid">{items}</div>
    </div>
    """, unsafe_allow_html=True)


def render_section_feedback(title, errors, enabled=True):
    if not enabled:
        return

    if errors:
        items = "".join(
            f"<li>{html.escape(error)}</li>"
            for error in errors
        )
        st.markdown(f"""
        <div class="section-feedback">
            <div class="section-feedback-title">{html.escape(title)}: что исправить</div>
            <ul>{items}</ul>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="section-feedback ok">
            <div class="section-feedback-title">{html.escape(title)}: раздел выглядит корректно</div>
            <ul><li>Можно переходить к следующему блоку или уточнить детали в примечании.</li></ul>
        </div>
        """, unsafe_allow_html=True)


def draft_safe_value(value):
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value

    if isinstance(value, tuple):
        return [draft_safe_value(item) for item in value]

    if isinstance(value, list):
        return [draft_safe_value(item) for item in value]

    if isinstance(value, dict):
        return {
            str(key): draft_safe_value(item)
            for key, item in value.items()
            if draft_safe_value(item) is not None
        }

    return None


DRAFT_SYSTEM_KEYS = {
    "draft_upload",
    "draft_expander_download",
    "floating_draft_download",
    "arm_clear_os_counts",
    "srv_fill_single_os",
    "srv_fill_remaining",
    "srv_clear_os_counts",
    "cached_report_bytes",
    "report_ready",
    "generation_active",
    "draft_link_ready",
    "draft_link_notice",
    "_draft_query_marker",
    "_arm_os_selection_mode",
}

DRAFT_FORBIDDEN_WIDGET_KEYS = {
    "draft_expander_download",
    "floating_draft_download",
    "arm_clear_os_counts",
    "srv_fill_single_os",
    "srv_fill_remaining",
    "srv_clear_os_counts",
}

SECURITY_CHECKBOX_KEYS = {
    "epp", "edr", "xdr", "mdr",
    "dlp", "mail_sec", "casb",
    "waf", "ddos", "nad", "ids", "nac", "ztna",
    "sast", "dast",
    "iam", "mfa", "pam",
    "siem", "soar",
    "vuln", "patch",
}


def coerce_draft_bool(value):
    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)):
        return value != 0

    if isinstance(value, str):
        return value.strip().lower() in {"true", "1", "yes", "y", "да", "вкл", "включено"}

    return False


def is_draft_system_key(key):
    key_text = str(key)
    action_markers = (
        "_clear_",
        "_fill_",
        "_download",
        "_generate",
        "_apply",
        "_button",
    )
    return (
        key_text in DRAFT_SYSTEM_KEYS
        or any(marker in key_text for marker in action_markers)
        or key_text.startswith("FormSubmitter:")
        or key_text.startswith("floating_")
    )


def is_forbidden_widget_state_key(key):
    key_text = str(key)
    return (
        key_text in DRAFT_FORBIDDEN_WIDGET_KEYS
        or key_text.startswith("FormSubmitter:")
    )


def clear_forbidden_widget_state():
    for key in list(st.session_state.keys()):
        if is_forbidden_widget_state_key(key):
            st.session_state.pop(key, None)


def collect_draft_state():
    draft_state = {}

    for key, value in st.session_state.items():
        if is_draft_system_key(key):
            continue

        safe_value = draft_safe_value(value)
        if safe_value is not None:
            draft_state[str(key)] = safe_value

    return {
        "schema": "khalil-audit-draft-v1",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "state": draft_state,
    }


def build_draft_download():
    draft_payload = collect_draft_state()
    draft_bytes = json.dumps(
        draft_payload,
        ensure_ascii=False,
        indent=2,
    ).encode("utf-8")
    draft_filename = f"khalil_audit_draft_{datetime.now():%Y%m%d_%H%M}.json"

    return draft_bytes, draft_filename


def encode_draft_token(payload):
    raw = json.dumps(
        payload,
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    compressed = zlib.compress(raw, level=9)
    return base64.urlsafe_b64encode(compressed).decode("ascii").rstrip("=")


def decode_draft_token(token):
    padding = "=" * (-len(token) % 4)
    compressed = base64.urlsafe_b64decode((token + padding).encode("ascii"))
    raw = zlib.decompress(compressed)
    return json.loads(raw.decode("utf-8"))


def apply_draft_state(payload):
    state = payload.get("state", payload)
    if not isinstance(state, dict):
        raise ValueError("Файл черновика не содержит блок state.")

    applied = 0

    for key, value in state.items():
        if is_draft_system_key(key):
            continue
        if key == "client_phone_code" and isinstance(value, list):
            value = tuple(value)
        if key in SECURITY_CHECKBOX_KEYS:
            value = coerce_draft_bool(value)
        st.session_state[key] = value
        applied += 1

    return applied


def restore_draft_from_query():
    token = st.query_params.get("draft")
    if not token:
        return

    if isinstance(token, list):
        token = token[0] if token else ""

    if not token:
        return

    marker = f"draft:{token[:24]}:{len(token)}"
    if st.session_state.get("_draft_query_marker") == marker:
        return

    try:
        payload = decode_draft_token(token)
        applied = apply_draft_state(payload)
        st.session_state["_draft_query_marker"] = marker
        st.session_state["draft_notice"] = (
            f"Черновик из ссылки применен: восстановлено полей {applied}."
        )
        st.query_params.clear()
        st.rerun()
    except Exception as exc:
        st.error(f"Не удалось применить черновик из ссылки: {exc}")


def render_floating_draft_save():
    draft_bytes, draft_filename = build_draft_download()
    draft_token = encode_draft_token(collect_draft_state())
    draft_text_json = json.dumps(draft_bytes.decode("utf-8"))
    draft_filename_json = json.dumps(draft_filename)
    draft_token_json = json.dumps(draft_token)
    can_share_json = "true" if len(draft_token) <= 6000 else "false"

    with st.container(key="floating_draft_save"):
        components.html(f"""
        <style>
        * {{
            box-sizing: border-box;
        }}

        .floating-draft-panel {{
            width: 100%;
            padding: 12px 14px 10px 14px;
            background: rgba(255, 255, 255, 0.97);
            border: 1px solid rgba(208, 213, 221, 0.92);
            border-radius: 22px;
            box-shadow: 0 18px 44px rgba(16, 24, 40, 0.16);
            backdrop-filter: blur(10px);
            font-family: "Inter", "Segoe UI", sans-serif;
        }}

        .floating-draft-actions {{
            display: grid;
            grid-template-columns: 1.35fr 1fr;
            gap: 12px;
            align-items: center;
        }}

        .floating-draft-action {{
            width: 100%;
            height: 52px;
            border-radius: 999px;
            border: 1px solid transparent;
            color: #ffffff;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            font-weight: 760;
            line-height: 1;
            margin: 0;
            padding: 0 18px;
            white-space: nowrap;
        }}

        .floating-draft-action.save {{
            background: #f97316;
            border-color: #ea580c;
            box-shadow: 0 12px 28px rgba(234, 88, 12, 0.22);
        }}

        .floating-draft-action.save:hover {{
            background: #ea580c;
            border-color: #c2410c;
        }}

        .floating-draft-action.share {{
            background: #0f766e;
            border-color: #0f766e;
            box-shadow: 0 12px 28px rgba(15, 118, 110, 0.18);
        }}

        .floating-draft-action.share:hover {{
            background: #0d5f59;
            border-color: #0d5f59;
        }}

        .floating-draft-action:disabled {{
            background: #98a2b3;
            border-color: #98a2b3;
            cursor: not-allowed;
            box-shadow: none;
        }}

        .floating-draft-hint {{
            color: #475467;
            font-size: 11px;
            line-height: 1.25;
            margin-top: 8px;
            text-align: center;
            text-shadow: 0 1px 2px rgba(255, 255, 255, 0.85);
        }}

        @media (max-width: 460px) {{
            .floating-draft-panel {{
                padding: 10px 10px 8px 10px;
                border-radius: 18px;
            }}

            .floating-draft-actions {{
                gap: 8px;
            }}

            .floating-draft-action {{
                height: 46px;
                font-size: 13px;
                padding: 0 10px;
            }}
        }}
        </style>
        <div class="floating-draft-panel">
            <div class="floating-draft-actions">
                <button class="floating-draft-action save" id="floating-save" type="button">
                    Сохранить черновик
                </button>
                <button class="floating-draft-action share" id="floating-share" type="button">
                    Поделиться
                </button>
            </div>
            <div class="floating-draft-hint" id="floating-draft-status">
                Сохраните JSON-файл или поделитесь ссылкой на заполненную анкету.
            </div>
        </div>
        <script>
        const draftText = {draft_text_json};
        const draftFilename = {draft_filename_json};
        const draftToken = {draft_token_json};
        const canShareDraft = {can_share_json};
        const saveButton = document.getElementById("floating-save");
        const shareButton = document.getElementById("floating-share");
        const status = document.getElementById("floating-draft-status");

        function getParentUrl() {{
            try {{
                return window.parent.location.href;
            }} catch (error) {{
                return document.referrer || window.location.href;
            }}
        }}

        function buildShareUrl() {{
            const url = new URL(getParentUrl());
            url.searchParams.set("draft", draftToken);
            return url.toString();
        }}

        function saveDraftFile() {{
            const blob = new Blob([draftText], {{
                type: "application/json;charset=utf-8"
            }});
            const href = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = href;
            link.download = draftFilename;
            document.body.appendChild(link);
            link.click();
            link.remove();
            window.setTimeout(() => URL.revokeObjectURL(href), 1000);
            status.textContent = "JSON-черновик скачан. Его можно загрузить в блоке «Черновик анкеты».";
        }}

        async function copyShareUrl(shareUrl) {{
            if (navigator.clipboard && window.isSecureContext) {{
                await navigator.clipboard.writeText(shareUrl);
                status.textContent = "Ссылка скопирована. Передайте ее коллеге.";
                return;
            }}

            window.prompt("Скопируйте ссылку на заполненную анкету", shareUrl);
            status.textContent = "Скопируйте ссылку из окна браузера.";
        }}

        async function shareDraft() {{
            if (!canShareDraft) {{
                status.textContent = "Анкета слишком большая для ссылки. Скачайте JSON-черновик.";
                return;
            }}

            const shareUrl = buildShareUrl();
            const shareData = {{
                title: "Khalil Audit System",
                text: "Заполненная анкета Khalil Audit System by Ivan Rudoy",
                url: shareUrl
            }};

            try {{
                if (navigator.share) {{
                    await navigator.share(shareData);
                    status.textContent = "Анкета готова к отправке.";
                    return;
                }}

                await copyShareUrl(shareUrl);
            }} catch (error) {{
                if (error && error.name === "AbortError") {{
                    status.textContent = "Отправка отменена.";
                    return;
                }}

                try {{
                    await copyShareUrl(shareUrl);
                }} catch (copyError) {{
                    window.prompt("Скопируйте ссылку на заполненную анкету", shareUrl);
                    status.textContent = "Скопируйте ссылку из окна браузера.";
                }}
            }}
        }}

        if (!canShareDraft) {{
            shareButton.disabled = true;
            shareButton.title = "Анкета слишком большая для ссылки";
        }}

        saveButton.addEventListener("click", saveDraftFile);
        shareButton.addEventListener("click", shareDraft);
        </script>
        """, height=96)


def render_draft_share_button(
    draft_token,
    label="Поделиться заполненной анкетой",
    status_text="Откроет меню отправки или скопирует ссылку.",
    height=82,
    compact=False,
):
    token_json = json.dumps(draft_token)
    label_json = json.dumps(label)
    status_json = json.dumps(status_text)
    status_markup = (
        '<div class="draft-share-status" id="draft-share-status"></div>'
        if status_text
        else '<div class="draft-share-status sr-only" id="draft-share-status"></div>'
    )
    border_radius = "999px" if compact else "8px"
    min_height = "42px" if compact else "38px"
    font_size = "14px" if compact else "14px"

    components.html(f"""
    <style>
    .draft-share-wrap {{
        font-family: "Inter", "Segoe UI", sans-serif;
        width: 100%;
    }}

    .draft-share-button {{
        width: 100%;
        min-height: {min_height};
        border: 1px solid #0f766e;
        border-radius: {border_radius};
        background: #0f766e;
        color: #ffffff;
        cursor: pointer;
        font-size: {font_size};
        font-weight: 650;
        line-height: 1.2;
    }}

    .draft-share-button:hover {{
        background: #0d5f59;
        border-color: #0d5f59;
    }}

    .draft-share-status {{
        color: #667085;
        font-size: 11px;
        line-height: 1.35;
        margin-top: 7px;
        text-align: center;
    }}

    .sr-only {{
        position: absolute;
        width: 1px;
        height: 1px;
        padding: 0;
        margin: -1px;
        overflow: hidden;
        clip: rect(0, 0, 0, 0);
        white-space: nowrap;
        border: 0;
    }}
    </style>
    <div class="draft-share-wrap">
        <button class="draft-share-button" id="draft-share-button" type="button">
        </button>
        {status_markup}
    </div>
    <script>
    const draftToken = {token_json};
    const buttonLabel = {label_json};
    const initialStatus = {status_json};
    const button = document.getElementById("draft-share-button");
    const status = document.getElementById("draft-share-status");
    button.textContent = buttonLabel;
    if (status) {{
        status.textContent = initialStatus;
    }}

    function getParentUrl() {{
        try {{
            return window.parent.location.href;
        }} catch (error) {{
            return document.referrer || window.location.href;
        }}
    }}

    function buildShareUrl() {{
        const url = new URL(getParentUrl());
        url.searchParams.set("draft", draftToken);
        return url.toString();
    }}

    async function copyShareUrl(shareUrl) {{
        if (navigator.clipboard && window.isSecureContext) {{
            await navigator.clipboard.writeText(shareUrl);
            status.textContent = "Ссылка скопирована. Передайте ее коллеге.";
            return;
        }}

        window.prompt("Скопируйте ссылку на заполненную анкету", shareUrl);
        status.textContent = "Скопируйте ссылку из окна браузера.";
    }}

    button.addEventListener("click", async () => {{
        const shareUrl = buildShareUrl();
        const shareData = {{
            title: "Khalil Audit System",
            text: "Заполненная анкета Khalil Audit System by Ivan Rudoy",
            url: shareUrl
        }};

        try {{
            if (navigator.share) {{
                await navigator.share(shareData);
                status.textContent = "Анкета готова к отправке.";
                return;
            }}

            await copyShareUrl(shareUrl);
        }} catch (error) {{
            if (error && error.name === "AbortError") {{
                status.textContent = "Отправка отменена.";
                return;
            }}

            try {{
                await copyShareUrl(shareUrl);
            }} catch (copyError) {{
                window.prompt("Скопируйте ссылку на заполненную анкету", shareUrl);
                status.textContent = "Скопируйте ссылку из окна браузера.";
            }}
        }}
    }});
    </script>
    """, height=height)


def render_draft_tools():
    if st.session_state.get("draft_notice"):
        st.success(st.session_state.pop("draft_notice"))

    with st.expander("Черновик анкеты: сохранить или продолжить заполнение", expanded=False):
        st.caption(
            "Сохраните JSON-черновик или создайте ссылку, если анкету нужно передать коллеге "
            "и продолжить заполнение на другом компьютере. Черновик содержит введенные ответы, "
            "поэтому передавайте его только доверенному получателю."
        )

        draft_bytes, draft_filename = build_draft_download()
        draft_token = encode_draft_token(collect_draft_state())

        col_save, col_link, col_load = st.columns(3)
        with col_save:
            st.download_button(
                "Скачать черновик JSON",
                data=draft_bytes,
                file_name=draft_filename,
                mime="application/json",
                use_container_width=True,
                key="draft_expander_download",
            )

        with col_link:
            if len(draft_token) > 6000:
                st.warning(
                    "Анкета уже слишком большая для надежной ссылки. "
                    "Скачайте JSON-черновик и передайте файл коллеге."
                )
            else:
                render_draft_share_button(draft_token)

        with col_load:
            uploaded_draft = st.file_uploader(
                "Загрузить черновик JSON",
                type=["json"],
                key="draft_upload",
                help="После применения страница перезагрузится и подставит сохраненные ответы.",
            )
            if uploaded_draft and st.button("Применить черновик", use_container_width=True):
                try:
                    payload = json.loads(uploaded_draft.getvalue().decode("utf-8-sig"))
                    applied = apply_draft_state(payload)
                    st.session_state["draft_notice"] = f"Черновик применен: восстановлено полей {applied}."
                    st.rerun()
                except Exception as exc:
                    st.error(f"Не удалось применить черновик: {exc}")


def render_generation_guard(active):
    action = "install" if active else "remove"
    components.html(f"""
    <script>
    const target = window.parent;
    const message = "Экспертный отчет формируется. Это может занять до 4 минут. Не закрывайте и не обновляйте страницу.";

    if (!target.__khalilBeforeUnloadHandler) {{
      target.__khalilBeforeUnloadHandler = (event) => {{
        event.preventDefault();
        event.returnValue = message;
        return message;
      }};
    }}

    if ("{action}" === "install" && !target.__khalilBeforeUnloadActive) {{
      target.addEventListener("beforeunload", target.__khalilBeforeUnloadHandler);
      target.__khalilBeforeUnloadActive = true;
    }}

    if ("{action}" === "remove" && target.__khalilBeforeUnloadActive) {{
      target.removeEventListener("beforeunload", target.__khalilBeforeUnloadHandler);
      target.__khalilBeforeUnloadActive = false;
    }}
    </script>
    """, height=0)


def render_generation_live_panel(stage_title, active_step=0):
    steps = [
        "Нормализация анкеты",
        "Сопоставление масштаба инфраструктуры",
        "Расчет зрелости ИТ/ИБ",
        "Формирование экспертного заключения",
        "Сборка XLSX и roadmap",
        "Отправка результата",
    ]
    step_items = []
    for idx, step in enumerate(steps):
        state_class = "done" if idx < active_step else "active" if idx == active_step else ""
        step_items.append(f'<li class="{state_class}">{html.escape(step)}</li>')

    components.html(f"""
    <style>
    .gen-panel {{
        font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        border: 1px solid #c7d7fe;
        border-left: 4px solid #13877c;
        border-radius: 10px;
        background: #ffffff;
        padding: 16px 18px;
        color: #101828;
        box-shadow: 0 10px 30px rgba(16, 24, 40, 0.08);
    }}
    .gen-head {{
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 10px;
    }}
    .gen-pulse {{
        width: 13px;
        height: 13px;
        border-radius: 50%;
        background: #13877c;
        box-shadow: 0 0 0 rgba(19, 135, 124, 0.5);
        animation: pulse 1.4s infinite;
        flex: 0 0 13px;
    }}
    @keyframes pulse {{
        0% {{ box-shadow: 0 0 0 0 rgba(19, 135, 124, 0.5); }}
        70% {{ box-shadow: 0 0 0 12px rgba(19, 135, 124, 0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(19, 135, 124, 0); }}
    }}
    .gen-title {{
        font-size: 17px;
        font-weight: 760;
    }}
    .gen-sub {{
        color: #475467;
        font-size: 13px;
        line-height: 1.45;
        margin-bottom: 12px;
    }}
    .gen-time {{
        display: inline-block;
        color: #0f5f59;
        background: #e6fffb;
        border: 1px solid #99f6e4;
        border-radius: 999px;
        padding: 5px 9px;
        font-size: 12px;
        font-weight: 760;
        margin-bottom: 12px;
    }}
    .gen-steps {{
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 8px;
        margin: 0;
        padding: 0;
        list-style: none;
    }}
    .gen-steps li {{
        border: 1px solid #e4e7ec;
        border-radius: 8px;
        padding: 8px 10px;
        color: #667085;
        background: #f8fafc;
        font-size: 12px;
        line-height: 1.3;
    }}
    .gen-steps li.done {{
        background: #f0fdf9;
        border-color: #99f6e4;
        color: #0f5f59;
        font-weight: 700;
    }}
    .gen-steps li.active {{
        background: #fff7ed;
        border-color: #fed7aa;
        color: #9a3412;
        font-weight: 760;
    }}
    </style>
    <div class="gen-panel">
        <div class="gen-head">
            <div class="gen-pulse"></div>
            <div class="gen-title">{html.escape(stage_title)}</div>
        </div>
        <div class="gen-sub">
            Формирование может занять до 4 минут. Если счетчик ниже идет, браузер не завис: отчет собирается на сервере.
        </div>
        <div class="gen-time">Прошло: <span id="gen-elapsed">00:00</span></div>
        <ul class="gen-steps">{''.join(step_items)}</ul>
    </div>
    <script>
    const startedAt = Date.now();
    const target = document.getElementById("gen-elapsed");
    setInterval(() => {{
        const total = Math.floor((Date.now() - startedAt) / 1000);
        const minutes = String(Math.floor(total / 60)).padStart(2, "0");
        const seconds = String(total % 60).padStart(2, "0");
        if (target) target.textContent = `${{minutes}}:${{seconds}}`;
    }}, 1000);
    </script>
    """, height=275)


inject_audit_design()

# Якорь для принудительного перехода в начало страницы
st.markdown("<div id='top'></div>", unsafe_allow_html=True)
# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = get_app_secret("TELEGRAM_TOKEN")
CHAT_ID = get_app_secret("TELEGRAM_CHAT_ID")

restore_draft_from_query()
clear_forbidden_widget_state()

render_app_header()
render_floating_draft_save()

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("Инструкция по заполнению", expanded=False):
    st.markdown("""
    ### Как пройти анкету

    1. **Начните с обязательного минимума:** заполните общую информацию о компании и блок **Конечные точки (АРМ)**. У заказчика почти всегда есть рабочие станции, поэтому этот блок не отключается и нужен для получения отчета.
    2. **Включайте только реальные разделы:** сеть, серверы, виртуализация, СХД, внутренние системы, ИБ и разработка открываются тумблерами. Если раздел включен, его вложенные обязательные поля нужно заполнить.
    3. **Смотрите навигатор слева:** красная точка означает, что раздел требует внимания, зеленая - раздел выглядит заполненным, серая - раздел отключен и не участвует в анкете.
    4. **Исправляйте по подсказкам на месте:** под каждым крупным блоком есть панель “что исправить”. Она показывает конкретные недостающие поля именно для этого раздела.
    5. **Используйте предварительную аналитику:** блок “Предварительная аналитика” показывает сводку аудита и быстрые улучшения еще до формирования финального XLSX. Полную версию можно раскрыть ниже.
    6. **Сохраняйте черновик при совместном заполнении:** скачайте JSON-черновик и передайте его коллеге. Он сможет загрузить файл в этом же блоке и продолжить заполнение с сохраненных ответов.
    7. **Не закрывайте страницу во время отчета:** формирование экспертного отчета может занять до 4 минут. Пока идет генерация, не обновляйте и не закрывайте вкладку.

    Поля “Примечание” необязательны, но помогают эксперту точнее описать контекст, ограничения и планы развития инфраструктуры.
    """)

render_draft_tools()

data = {}
client_info = {}
validation_errors = []
score = 0
# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
render_section_marker(
    "01 / КОМПАНИЯ",
    "Общая информация",
    "Контекст компании, отрасль и контактные данные для экспертного отчета.",
    anchor_id="section-company"
)
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город*", key="client_city", help="Укажите город фактического нахождения головного офиса.")
    industry_options = ["Финтех / Банки", "Ритейл / E-commerce", "Производство", "IT / Разработка", "Госсектор", "Другое"]
    selected_ind = st.selectbox(
        "Сфера деятельности компании*",
        [""] + industry_options,
        format_func=lambda x: "Выберите сферу..." if x == "" else x,
        key="client_industry_select",
        help="Отрасль влияет на профиль угроз и регуляторные требования."
    )

    if selected_ind == "Другое":
        industry = st.text_input("Укажите вашу сферу деятельности*", key="client_industry_other", help="Введите отрасль вручную")
    else:
        industry = selected_ind
    
    client_info['Сфера деятельности'] = industry
    client_info['Наименование компании'] = st.text_input("Наименование компании*", key="client_company_name", help="Официальное или сокращенное название юрлица.")

    site_input = st.text_input("Сайт компании*", key="site_field", placeholder="example.kz", help="Можно указать example.kz или https://example.kz - система приведет адрес к домену.")
    clean_domain = normalize_site_domain(site_input)
    client_info['Сайт компании'] = clean_domain
    if site_input and clean_domain != site_input.strip().lower():
        st.caption(f"Будет использован домен: {clean_domain}")

    custom_email_mode = st.checkbox("Email отличается от сайта", key="client_custom_email_mode", help="Отметьте, если корпоративная почта находится на другом домене.")

    if custom_email_mode:
        custom_email = st.text_input("Email контактного лица*", key="client_email_custom", help="Личный корпоративный email для отправки результатов.")
        client_info['Email'] = normalize_email(custom_email)
    else:
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин)*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre", help="Только часть адреса до символа @")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = normalize_email(f"{email_prefix}@{clean_domain}") if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*", key="client_contact_name", help="С кем наш эксперт сможет обсудить детали отчета.")
    client_info['Должность'] = st.text_input("Должность*", key="client_contact_role", help="Например: ИТ-Директор, Системный администратор, CEO.")
    
    st.write("Контактный телефон*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996"),
        ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90"), ("🇦🇿 +994", "+994")
    ]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed", key="client_phone_code")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed", key="client_phone_number", help="Телефон для оперативной связи.")
    normalized_phone = normalize_phone_number(phone_num)
    if phone_num and normalized_phone != phone_num.strip():
        st.caption(f"Номер будет записан как: {selected_code[1]} {normalized_phone}")
    client_info['Контактный телефон'] = f"{selected_code[1]} {normalized_phone}" if phone_num else ""

if not all([client_info.get('Город'), client_info.get('Наименование компании'), client_info.get('Сфера деятельности'), client_info.get('Сайт компании'), client_info.get('Email'), client_info.get('ФИО контактного лица'), client_info.get('Должность'), phone_num]):
    validation_errors.append("Заполните все обязательные поля в блоке 'Общая информация'")
elif not is_valid_domain(client_info.get('Сайт компании')):
    validation_errors.append("Укажите корректный сайт компании")
elif not is_valid_email(client_info.get('Email')):
    validation_errors.append("Укажите корректный email контактного лица")
elif phone_num and len(re.sub(r"\D", "", phone_num)) < 7:
    validation_errors.append("Укажите корректный контактный телефон")

render_section_feedback(
    "Общая информация",
    get_section_errors(validation_errors, "общая информация", "сайт", "email", "телефон")
)

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
render_section_marker(
    "02 / ИНФРАСТРУКТУРА",
    "Информационные технологии",
    "АРМ, сеть, серверы, хранение данных и внутренние бизнес-системы."
)

render_anchor("section-endpoints")
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1, key="total_arm", help="Общее число ПК, ноутбуков и тонких клиентов в организации.")
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"], key="selected_os_arm", help="Выберите все типы операционных систем, используемых сотрудниками.")

sum_os_arm = 0
if selected_os_arm:
    arm_keys = [f"arm_{os_item}" for os_item in selected_os_arm]
    if len(selected_os_arm) == 1:
        only_os = selected_os_arm[0]
        st.session_state[f"arm_{only_os}"] = total_arm
        st.session_state["_arm_os_selection_mode"] = "single"
        st.caption(f"Выбрана одна ОС, поэтому количество автоматически равно общему числу АРМ: {total_arm}.")
    else:
        if st.session_state.get("_arm_os_selection_mode") == "single":
            for key in arm_keys:
                st.session_state[key] = 0
            st.session_state["_arm_os_selection_mode"] = "multi"
            st.rerun()

        st.session_state["_arm_os_selection_mode"] = "multi"

    if len(selected_os_arm) > 1 and total_arm > 0:
        if st.button("Очистить распределение ОС", use_container_width=True, key="arm_clear_os_counts"):
            for key in arm_keys:
                st.session_state[key] = 0
            st.rerun()
else:
    st.session_state["_arm_os_selection_mode"] = "empty"

for os_item in selected_os_arm:
    val = st.number_input(f"Кол-во на {os_item}", min_value=0, step=1, key=f"arm_{os_item}", help=f"Укажите точное или примерное число устройств с {os_item}.")
    data[f"ОС АРМ ({os_item})"] = val
    sum_os_arm += val

data['1.1. Примечание'] = st.text_area("Примечание к разделу 1.1", placeholder="Напр.: планируем обновление Windows 10 до 11 в Q3", key="note_1_1")

if total_arm <= 0:
    validation_errors.append("Укажите количество АРМ")
elif not selected_os_arm:
    validation_errors.append("Выберите ОС на АРМ")
elif sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка: Сумма по ОС ({sum_os_arm}) должна быть равна общему количеству АРМ ({total_arm}).")
    validation_errors.append("Несовпадение количества АРМ и ОС")

render_section_feedback(
    "Конечные точки",
    get_section_errors(validation_errors, "арм", "ос на арм")
)

st.write("---")
render_anchor("section-network")
st.subheader("1.2. Сетевая инфраструктура")
main_speed, back_speed, ap_cnt = 0, 0, 0
selected_routing = []
ngfw_vendor = "Нет"
wifi_enabled = False
wifi_ctrl_enabled = False
net_active = st.toggle("Своя сетевая инфраструктура", key="net_toggle", help="Активируйте, если организация самостоятельно управляет сетевым оборудованием.")

if net_active:
    net_types = ["Оптика", "RJ45 (Ethernet)", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    routing_types = ["Статическая", "RIP", "OSPF", "EIGRP", "BGP", "IS-IS"]
    
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        st.write("Основной канал")
        main_net_kwargs = {"key": "main_net_type", "help": "Технология подключения основного интернет-канала."}
        if "main_net_type" not in st.session_state:
            main_net_kwargs["index"] = 7
        main_type = st.selectbox("Тип (основной)", net_types, **main_net_kwargs)
        main_speed = st.number_input("Скорость основного (Mbit/s)", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbit/s)"
    with col_net2:
        st.write("Резервный канал")
        back_net_kwargs = {"key": "back_net_type", "help": "Наличие и тип независимого резервного канала."}
        if "back_net_type" not in st.session_state:
            back_net_kwargs["index"] = 7
        back_type = st.selectbox("Тип (резервный)", net_types, **back_net_kwargs)
        back_speed = st.number_input("Скорость резервного (Mbit/s)", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbit/s)"

    st.write("Логика сети")
    selected_routing = st.multiselect("Тип маршрутизации*", routing_types, key="routing_sel", help="Протоколы динамической маршрутизации, используемые в сети.")
    data['1.2.3. Маршрутизация'] = ", ".join(selected_routing)
    if not selected_routing:
        validation_errors.append("Выберите тип маршрутизации")

    st.write("Активное сетевое оборудование")
    c_net1, c_net2, c_net3 = st.columns(3)
    with c_net1:
        if st.checkbox("Маршрутизаторы", key="router_chk", help="Устройства для связи разных сетей и выхода в интернет."):
            r_count = st.number_input("Кол-во маршрутизаторов", min_value=0, step=1, key="router_cnt")
            data['1.2.4. Маршрутизаторы'] = f"Да ({r_count} шт)"
            if r_count == 0: validation_errors.append("Укажите количество маршрутизаторов")
    with c_net2:
        if st.checkbox("Коммутаторы L2", key="swl2_chk", help="Управляемые или неуправляемые коммутаторы уровня доступа."):
            sw2_count = st.number_input("Кол-во коммутаторов L2", min_value=0, step=1, key="swl2_cnt")
            data['1.2.5. Коммутаторы L2'] = f"Да ({sw2_count} шт)"
            if sw2_count == 0: validation_errors.append("Укажите количество коммутаторов L2")
    with c_net3:
        if st.checkbox("Коммутаторы L3", key="swl3_chk", help="Коммутаторы с функциями маршрутизации (ядро или агрегация)."):
            sw3_count = st.number_input("Кол-во коммутаторов L3", min_value=0, step=1, key="swl3_cnt")
            data['1.2.6. Коммутаторы L3'] = f"Да ({sw3_count} шт)"
            if sw3_count == 0: validation_errors.append("Укажите количество коммутаторов L3")

    st.write("Уровни сети")
    l_col1, l_col2, l_col3 = st.columns(3)
    with l_col1:
        if st.checkbox("Ядро (Core)", key="net_core", help="Центральная часть сети, обеспечивающая максимальную скорость."):
            core_v = st.text_input("Основной производитель (Core)", key="core_vendor", help="Например: Cisco, Huawei, Juniper, MikroTik.")
            data['Уровень сети Ядро'] = core_v
            if not core_v: validation_errors.append("Укажите производителя Core-уровня")
    with l_col2:
        if st.checkbox("Уровень распределения", key="net_dist", help="Связующее звено между ядром и уровнем доступа."):
            dist_v = st.text_input("Основной производитель (Dist)", key="dist_vendor")
            data['Уровень сети Распределение'] = dist_v
            if not dist_v: validation_errors.append("Укажите производителя уровня распределения")
    with l_col3:
        if st.checkbox("Уровень доступа", key="net_acc", help="Уровень, к которому подключаются конечные пользователи."):
            acc_v = st.text_input("Основной производитель (Access)", key="acc_vendor")
            data['Уровень сети Доступ'] = acc_v
            if not acc_v: validation_errors.append("Укажите производителя уровня доступа")

    wifi_enabled = st.checkbox("Wi-Fi", key="wifi_toggle", help="Наличие корпоративной беспроводной сети.")
    if wifi_enabled:
        w_col1, w_col2, w_col3 = st.columns(3)
        with w_col1:
            wifi_ctrl_enabled = st.checkbox("Контроллер", key="wifi_ctrl", help="Централизованное управление точками доступа (аппаратное или программное).")
            if wifi_ctrl_enabled:
                wc_v = st.text_input("Производитель/модель контроллера", key="wc_vendor")
                data['Wi-Fi Контроллер'] = wc_v
                if not wc_v: validation_errors.append("Укажите модель Wi-Fi контроллера")
            else:
                data['Wi-Fi Контроллер'] = "Нет"
        with w_col2:
            ap_cnt = st.number_input("Количество точек доступа (шт)", min_value=0, step=1, key="ap_cnt", help="Общее число активных Wi-Fi точек.")
            data['Wi-Fi Точки доступа'] = ap_cnt
            if ap_cnt == 0: validation_errors.append("Укажите количество точек доступа Wi-Fi")
        with w_col3:
            wf_types = ["Wi-Fi 6/6E (802.11ax)", "Wi-Fi 5 (802.11ac)", "Wi-Fi 4 (802.11n)", "Другое"]
            data['Wi-Fi Тип'] = st.selectbox("Тип Wi-Fi", wf_types, key="wf_type_sel", help="Преимущественный стандарт беспроводной связи.")

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk", help="Многофункциональные шлюзы безопасности (FortiGate, UserGate, CheckPoint и т.д.)."):
        ngfw_vendor = st.text_input("Производитель (NGFW)", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor if ngfw_vendor else 'не указан'})"
        if not ngfw_vendor: validation_errors.append("Укажите производителя NGFW")
        score += 20
    
    data['1.2. Примечание'] = st.text_area("Примечание к разделу 1.2", placeholder="Особенности топологии сети...", key="note_1_2")

render_section_feedback(
    "Сетевая инфраструктура",
    get_section_errors(
        validation_errors,
        "маршрутизац",
        "маршрутизатор",
        "коммутатор",
        "core",
        "распредел",
        "доступ",
        "wi-fi",
        "ngfw"
    ),
    enabled=net_active
)

st.write("---")
render_anchor("section-servers")
st.subheader("1.3. Серверы и Виртуализация")
sum_os_srv = 0
phys_count = 0
virt_count = 0
server_active = st.toggle("Серверы и виртуализация", key="server_toggle", help="Включите, если у заказчика есть физические серверы, виртуальные машины или платформы виртуализации.")
v_n_b = "Нет"
selected_os_srv = []
selected_virt_sys = []

if server_active:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        phys_count = st.number_input("Количество физических серверов", min_value=0, step=1, key="phys_srv", help="Количество 'железных' серверов в серверной или ЦОД.")
        data['1.3.1. Физические серверы'] = phys_count
    with col_s2:
        virt_count = st.number_input("Количество виртуальных серверов", min_value=0, step=1, key="virt_srv", help="Суммарное количество виртуальных машин (VM).")
        data['1.3.2. Виртуальные серверы'] = virt_count

    if phys_count == 0 and virt_count == 0:
        validation_errors.append("Укажите количество физических или виртуальных серверов")

    s_os_list = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Unix", "Другое"]
    selected_os_srv = st.multiselect("Выберите ОС серверов", s_os_list, key="ms_srv_list", help="Операционные системы, установленные на серверах.")
    if selected_os_srv:
        srv_os_keys = [f"fsrv_{os_s}" for os_s in selected_os_srv]
        current_srv_os_total = sum(widget_int(key) for key in srv_os_keys)
        server_os_target = max(phys_count, virt_count)
        server_os_remaining = server_os_target - current_srv_os_total

        if server_os_target > 0:
            srv_quick_cols = st.columns(2)
            if len(selected_os_srv) == 1:
                only_srv_os = selected_os_srv[0]
                with srv_quick_cols[0]:
                    if st.button(f"Все {server_os_target} на {only_srv_os}", use_container_width=True, key="srv_fill_single_os"):
                        st.session_state[f"fsrv_{only_srv_os}"] = server_os_target
                        st.rerun()
            elif server_os_remaining > 0:
                target_srv_os = selected_os_srv[-1]
                with srv_quick_cols[0]:
                    if st.button(f"Добавить остаток {server_os_remaining} к {target_srv_os}", use_container_width=True, key="srv_fill_remaining"):
                        st.session_state[f"fsrv_{target_srv_os}"] = widget_int(f"fsrv_{target_srv_os}") + server_os_remaining
                        st.rerun()
            with srv_quick_cols[1]:
                if st.button("Очистить распределение серверных ОС", use_container_width=True, key="srv_clear_os_counts"):
                    for key in srv_os_keys:
                        st.session_state[key] = 0
                    st.rerun()

        for os_s in selected_os_srv:
            val_os = st.number_input(f"Кол-во на {os_s}", min_value=0, key=f"fsrv_{os_s}")
            data[f"ОС Сервера ({os_s})"] = val_os
            sum_os_srv += val_os

    if virt_count > 0 and sum_os_srv < virt_count:
        st.warning(f"⚠️ Ошибка: Количество ОС ({sum_os_srv}) должно быть больше или равно количеству виртуальных серверов ({virt_count}).")
        validation_errors.append("Недостаточное количество ОС для серверов")

    selected_virt_sys = st.multiselect("Выберите системы виртуализации", ["VMware", "Hyper-V", "Proxmox", "KVM", "Другое", "Нет"], key="virt_sys_list", help="Технологии управления виртуальной инфраструктурой.")
    if selected_virt_sys and "Нет" not in selected_virt_sys:
        for v_sys in selected_virt_sys:
            v_h_cnt = st.number_input(f"Количество хостов {v_sys}", min_value=0, step=1, key=f"fv_cnt_{v_sys}", help=f"Сколько физических серверов (нод) в кластере {v_sys}?")
            data[f"Система виртуализации ({v_sys})"] = v_h_cnt
            if v_h_cnt == 0: validation_errors.append(f"Укажите количество хостов для {v_sys}")

    if st.checkbox("Резервное копирование", key="ib_backup", help="Наличие специализированного ПО для бэкапа (Veeam, Commvault, Veritas и т.д.)."):
        v_n_b = st.text_input("Вендор Резервного копирования", key="vn_backup", help="Укажите название используемого продукта.")
        data["Резервное копирование"] = v_n_b
        if not v_n_b: validation_errors.append("Укажите вендора резервного копирования")
        score += 20

    data['1.3. Примечание'] = st.text_area("Примечание к разделу 1.3", placeholder="Специфика серверного парка...", key="note_1_3")

render_section_feedback(
    "Серверы и виртуализация",
    get_section_errors(
        validation_errors,
        "сервер",
        "хост",
        "резервного копирования"
    ),
    enabled=server_active
)

st.write("---")
render_anchor("section-storage")
st.subheader("1.4. Системы хранения данных (СХД)")
st_media_sel = []
cnt_hdd = 0
cnt_ssd = 0
raid_selected = []
storage_active = st.toggle("Есть собственная СХД", key="storage_toggle")
if storage_active:
    st_media_sel = st.multiselect(
        "Типы носителей",
        ["HDD (NL-SAS / SATA)", "SSD (SATA / SAS)", "NVMe", "SCM"],
        key="st_media"
    )
    data['1.4.1. Типы носителей'] = ", ".join(st_media_sel) if st_media_sel else "Не указано"

    col_pct1, col_pct2 = st.columns(2)
    with col_pct1:
        cnt_hdd = st.number_input("Количество дисков HDD", min_value=0, step=1, key="cnt_hdd")
        data['1.4.2. Кол-во HDD'] = cnt_hdd
    with col_pct2:
        cnt_ssd = st.number_input("Количество дисков SSD", min_value=0, step=1, key="cnt_ssd")
        data['1.4.3. Кол-во SSD'] = cnt_ssd

    if st_media_sel and (cnt_hdd + cnt_ssd == 0):
        st.info("ℹ️ Укажите количество дисков для СХД.")
        validation_errors.append("Не указано количество дисков СХД")

    col_chk1, col_chk2 = st.columns(2)
    with col_chk1:
        hybrid = st.checkbox("Используется гибридная СХД", key="hybrid_st")
        data['1.4.4. Гибридная СХД'] = "Да" if hybrid else "Нет"
    with col_chk2:
        allflash = st.checkbox("Есть All-Flash массивы", key="allflash_st")
        data['1.4.5. All-Flash'] = "Да" if allflash else "Нет"
        if allflash: score += 5

    raid_selected = st.multiselect(
        "Используемые RAID-группы",
        ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "RAID 50", "RAID 60", "JBOD"],
        key="raid_list"
    )
    data['1.4.6. RAID-группы'] = ", ".join(raid_selected) if raid_selected else "Не указано"

    if not raid_selected:
        validation_errors.append("Не указаны RAID-группы СХД")
    if "RAID 0" in raid_selected or "JBOD" in raid_selected:
        score -= 10

    data['1.4. Примечание'] = st.text_area("Примечание к разделу 1.4", placeholder="SAN/NAS, replication, snapshot, DR-site, tiering и т.д.", key="note_1_4")

render_section_feedback(
    "Системы хранения данных",
    get_section_errors(validation_errors, "схд", "raid"),
    enabled=storage_active
)

st.write("---")
render_anchor("section-systems")
st.subheader("1.5. Внутренние Информационные системы")
is_active = st.toggle("ИС организации", key="is_toggle", help="Бизнес-приложения и корпоративные сервисы.")
if is_active:
    is_types = {
        "ERP": "erp", "CRM": "crm", "HelpDesk/ServiceDesk": "sd", 
        "СЭД (Документооборот)": "sed", "HRM (Кадры)": "hrm", 
        "BI (Аналитика)": "bi", "WMS (Склад)": "wms", "Учет (Бухгалтерия)": "acc"
    }
    
    m_opts = ["Exchange (On-Prem)", "Lotus", "Microsoft 365", "Google Workspace", "Собственный", "Нет"]
    m_sys = st.selectbox("Почтовая система", m_opts, key="mail_system", help="Где физически и логически располагается ваша электронная почта.")
    
    if m_sys in ["Exchange (On-Prem)", "Lotus"]:
        m_ver = st.text_input(f"Версия {m_sys}*", key="mail_version_input", help="Например: 2016 CU23 или v9.0.1.")
        data['1.5.1. Почтовая система'] = f"{m_sys} (v.{m_ver})"
        if not m_ver: validation_errors.append(f"Укажите версию {m_sys}")
    else:
        data['1.5.1. Почтовая система'] = m_sys

    for label, ks in is_types.items():
        if st.checkbox(label, key=f"is_chk_{ks}"):
            c_is1, c_is2 = st.columns(2)
            with c_is1:
                name_is = st.text_input(f"Название продукта {label}*", key=f"name_{ks}")
            with c_is2:
                ver_is = st.text_input(f"Версия {label}*", key=f"ver_{ks}")
            data[f"ИС {label}"] = f"{name_is} (v.{ver_is})"
            if not name_is or not ver_is:
                validation_errors.append(f"Укажите название и версию для {label}")
    
    data['1.5. Примечание'] = st.text_area("Примечание к разделу 1.5", placeholder="Дополнительные ИС...", key="note_1_5")

render_section_feedback(
    "Внутренние информационные системы",
    get_section_errors(validation_errors, "версию", "название и версию"),
    enabled=is_active
)

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
render_section_marker(
    "03 / БЕЗОПАСНОСТЬ",
    "Информационная безопасность",
    "Текущие средства защиты, мониторинг, доступ, приложения и управление уязвимостями.",
    anchor_id="section-security"
)

enable_security_kwargs = {"key": "enable_security"}
if "enable_security" not in st.session_state:
    enable_security_kwargs["value"] = False
enable_security = st.toggle("Включить блок ИБ", **enable_security_kwargs)

# Инициализация переменных ИБ
epp, epp_v, edr, edr_v, xdr, xdr_v, mdr, mdr_v = False, "", False, "", False, "", False, ""
dlp, dlp_v, mail_sec, mail_v, casb, casb_v = False, "", False, "", False, ""
waf, waf_v, ddos, ddos_v, ids, ids_v, nac, nac_v, ztna, ztna_v = False, "", False, "", False, "", False, "", False, ""
sast, sast_v, dast, dast_v = False, "", False, ""
iam, iam_v, mfa, mfa_v, pam, pam_v = False, "", False, "", False, ""
siem, siem_v, soar, soar_v = False, "", False, ""
vuln, vuln_v, patch, patch_v, nad, nad_v = False, "", False, "", False, ""

def security_product(label, checkbox_key, vendor_label, vendor_key):
    enabled = st.checkbox(label, key=checkbox_key) is True
    if enabled:
        return enabled, st.text_input(vendor_label, key=vendor_key)
    st.session_state.pop(vendor_key, None)
    return enabled, ""

if enable_security:
    errors = []

    # =========================
    # Защита конечных устройств
    # =========================
    st.markdown("#### Защита конечных устройств")
    col1, col2 = st.columns(2)
    with col1:
        epp, epp_v = security_product("EPP (антивирусная защита)", "epp", "Производитель EPP", "epp_v")
        data['Блок 2. EPP'] = epp_v if epp else "Нет"

        edr, edr_v = security_product("EDR (обнаружение и реагирование)", "edr", "Производитель EDR", "edr_v")
        data['Блок 2. EDR'] = edr_v if edr else "Нет"
    with col2:
        xdr, xdr_v = security_product("XDR (расширенная защита)", "xdr", "Производитель XDR", "xdr_v")
        data['Блок 2. XDR'] = xdr_v if xdr else "Нет"

        mdr, mdr_v = security_product("MDR (внешний мониторинг)", "mdr", "Провайдер MDR", "mdr_v")
        data['Блок 2. MDR'] = mdr_v if mdr else "Нет"

    # =========================
    # Защита данных
    # =========================
    st.markdown("#### Защита данных")
    col1, col2 = st.columns(2)
    with col1:
        dlp, dlp_v = security_product("DLP (предотвращение утечек)", "dlp", "Производитель DLP", "dlp_v")
        data['Блок 2. DLP'] = dlp_v if dlp else "Нет"

        mail_sec, mail_v = security_product("Mail Security (защита почты)", "mail_sec", "Производитель Mail Security", "mail_v")
        data['Блок 2. Mail Security'] = mail_v if mail_sec else "Нет"
    with col2:
        casb, casb_v = security_product("CASB (контроль облаков)", "casb", "Производитель CASB", "casb_v")
        data['Блок 2. CASB'] = casb_v if casb else "Нет"

    # =========================
    # Сетевая безопасность
    # =========================
    st.markdown("#### Сетевая безопасность")
    col1, col2 = st.columns(2)
    with col1:
        waf, waf_v = security_product("WAF (защита веб-приложений)", "waf", "Производитель WAF", "waf_v")
        data['Блок 2. WAF'] = waf_v if waf else "Нет"

        ddos, ddos_v = security_product("Anti-DDoS (защита от атак)", "ddos", "Производитель Anti-DDoS", "ddos_v")
        data['Блок 2. Anti-DDoS'] = ddos_v if ddos else "Нет"

        nad, nad_v = security_product("NAD (Network Attack Discovery)", "nad", "Производитель NAD", "nad_v")
        data['Блок 2. NAD'] = nad_v if nad else "Нет"
    with col2:
        ids, ids_v = security_product("IDS/IPS (сетевые атаки)", "ids", "Производитель IDS/IPS", "ids_v")
        data['Блок 2. IDS/IPS'] = ids_v if ids else "Нет"

        nac, nac_v = security_product("NAC (контроль доступа)", "nac", "Производитель NAC", "nac_v")
        data['Блок 2. NAC'] = nac_v if nac else "Нет"

        ztna, ztna_v = security_product("ZTNA (Zero Trust доступ)", "ztna", "Производитель ZTNA", "ztna_v")
        data['Блок 2. ZTNA'] = ztna_v if ztna else "Нет"

    # =========================
    # Безопасность приложений
    # =========================
    st.markdown("#### Безопасность приложений")
    col1, col2 = st.columns(2)
    with col1:
        sast, sast_v = security_product("SAST (анализ кода)", "sast", "Производитель SAST", "sast_v")
        data['Блок 2. SAST'] = sast_v if sast else "Нет"
    with col2:
        dast, dast_v = security_product("DAST (тестирование приложений)", "dast", "Производитель DAST", "dast_v")
        data['Блок 2. DAST'] = dast_v if dast else "Нет"

    # =========================
    # Управление доступом
    # =========================
    st.markdown("#### Управление доступом")
    col1, col2 = st.columns(2)
    with col1:
        iam, iam_v = security_product("IAM (учетные записи)", "iam", "Производитель IAM", "iam_v")
        data['Блок 2. IAM'] = iam_v if iam else "Нет"

        mfa, mfa_v = security_product("MFA (многофакторная аутентификация)", "mfa", "Производитель MFA", "mfa_v")
        data['Блок 2. MFA'] = mfa_v if mfa else "Нет"
    with col2:
        pam, pam_v = security_product("PAM (привилегированный доступ)", "pam", "Производитель PAM", "pam_v")
        data['Блок 2. PAM'] = pam_v if pam else "Нет"

    # =========================
    # SOC
    # =========================
    st.markdown("#### Мониторинг и реагирование")
    col1, col2 = st.columns(2)
    with col1:
        siem, siem_v = security_product("SIEM (мониторинг событий)", "siem", "Производитель SIEM", "siem_v")
        data['Блок 2. SIEM'] = siem_v if siem else "Нет"
    with col2:
        soar, soar_v = security_product("SOAR (автоматизация)", "soar", "Производитель SOAR", "soar_v")
        data['Блок 2. SOAR'] = soar_v if soar else "Нет"

    # =========================
    # ДОПОЛНИТЕЛЬНО
    # =========================
    st.markdown("#### Дополнительно")
    col1, col2 = st.columns(2)
    with col1:
        vuln, vuln_v = security_product("Сканер уязвимостей", "vuln", "Производитель сканера", "vuln_v")
        data['Блок 2. Сканер уязвимостей'] = vuln_v if vuln else "Нет"
    with col2:
        patch, patch_v = security_product("Patch Management (управление обновлениями)", "patch", "Производитель Patch Management", "patch_v")
        data['Блок 2. Patch Management'] = patch_v if patch else "Нет"

    # Валидация
    ib_items = [
        ("EPP", epp, epp_v), ("EDR", edr, edr_v), ("XDR", xdr, xdr_v),
        ("MDR", mdr, mdr_v), ("DLP", dlp, dlp_v), ("Mail Security", mail_sec, mail_v),
        ("CASB", casb, casb_v), ("WAF", waf, waf_v), ("Anti-DDoS", ddos, ddos_v),
        ("IDS/IPS", ids, ids_v), ("NAC", nac, nac_v), ("ZTNA", ztna, ztna_v),
        ("SAST", sast, sast_v), ("DAST", dast, dast_v),
        ("IAM", iam, iam_v), ("MFA", mfa, mfa_v), ("PAM", pam, pam_v),
        ("SIEM", siem, siem_v), ("SOAR", soar, soar_v),
        ("Сканер уязвимостей", vuln, vuln_v),
        ("Patch Management", patch, patch_v),
        ("NAD", nad, nad_v)
    ]
    for name, enabled, vendor in ib_items:
        if enabled and not vendor:
            errors.append(f"Не указан производитель: {name}")

    if errors:
        st.error("Заполните обязательные поля в блоке ИБ:")
        for e in errors:
            st.write(f"- {e}")
        validation_errors.extend(errors)

render_section_feedback(
    "Информационная безопасность",
    get_section_errors(validation_errors, "производитель", "иб"),
    enabled=enable_security
)

# --- БЛОК 3: WEB-РЕСУРСЫ ---
render_section_marker(
    "04 / ЦИФРОВАЯ ПОВЕРХНОСТЬ",
    "Веб-ресурсы",
    "Публичная поверхность, фронтенд-стек и особенности размещения.",
    anchor_id="section-web"
)
web_active = st.toggle("Веб-ресурсы", key="web_toggle")
if web_active:
    data['3.1. Хостинг'] = st.selectbox("Хостинг", ["Собственный ЦОД", "Облако KZ", "Облако Global"], key="web_hosting")
    data['3.2. Frontend'] = st.multiselect("Frontend серверы", ["Nginx", "Apache", "IIS", "LiteSpeed", "Cloudflare"], key="web_frontend")
    data['Примечание (Web)'] = st.text_area("Примечания по Web", placeholder="Стек...", key="note_web")

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
render_section_marker(
    "05 / РАЗРАБОТКА",
    "Разработка",
    "Команда разработки, языки, CI/CD и дополнительный технологический контекст.",
    anchor_id="section-dev"
)
dev_count = 0
sel_langs = []
cicd_active = False
dev_active = st.toggle("Разработка", key="dev_toggle")
if dev_active:
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        dev_count = st.number_input("Кол-во разработчиков*", min_value=0, key="dev_cnt_f")
        data['4.1. Разработчики'] = dev_count
        cicd_active = st.checkbox("Используется CI/CD", key="cicd_f")
        data['4.2. CICD'] = "Да" if cicd_active else "Нет"
        if dev_count == 0: validation_errors.append("Укажите количество разработчиков")
    with col_d2:
        lang_list = ["Python", "JavaScript/TypeScript", "Java", "C# / .NET", "PHP", "Go", "C++", "Swift/Kotlin", "Другое"]
        sel_langs = st.multiselect("Языки программирования*", lang_list, key="langs_f")
        if not sel_langs:
            validation_errors.append("Выберите языки разработки")
            data['4.3. Языки разработки'] = "Не указаны"
        elif "Другое" in sel_langs:
            other_l = st.text_input("Укажите другие языки", key="other_langs_f")
            data['4.3. Языки разработки'] = f"{', '.join([l for l in sel_langs if l != 'Другое'])}, {other_l}"
        else:
            data['4.3. Языки разработки'] = ", ".join(sel_langs)
    data['Блок 4. Примечание'] = st.text_area("Примечание к разделу Разработка", placeholder="Стек, фреймворки...", key="note_dev")

render_section_feedback(
    "Разработка",
    get_section_errors(validation_errors, "разработчик", "языки разработки"),
    enabled=dev_active
)
#----Подготовка---
def get_maturity_level(score):
    if score <= 20:
        return "Начальный", "🔴"
    elif score <= 40:
        return "Базовый", "🟠"
    elif score <= 60:
        return "Развивающийся", "🟡"
    elif score <= 80:
        return "Управляемый", "🟢"
    else:
        return "Оптимальный", "🔵"


def calculate_weighted_security_score(enabled, controls):
    if not enabled:
        return 0

    total = 0
    for is_enabled, vendor, weight in controls:
        if is_enabled and str(vendor).strip():
            total += weight

    return min(100, total)


def calculate_it_maturity_score(
    total_arm,
    selected_os_arm,
    sum_os_arm,
    net_active,
    main_speed,
    back_speed,
    selected_routing,
    ap_cnt,
    ngfw_vendor,
    server_active,
    phys_count,
    virt_count,
    selected_virt_sys,
    backup_vendor,
    storage_active,
    st_media_sel,
    cnt_hdd,
    cnt_ssd,
    raid_selected,
    systems_active,
    web_active,
    dev_active,
    dev_count,
    sel_langs,
    cicd_active,
):
    maturity_score = 0

    if total_arm > 0:
        maturity_score += 6
        if selected_os_arm and sum_os_arm == total_arm:
            maturity_score += 5
        if total_arm >= 10:
            maturity_score += 4
        if total_arm >= 50:
            maturity_score += 5
        if total_arm >= 250:
            maturity_score += 5

    if net_active:
        if main_speed > 0:
            maturity_score += 4
        if main_speed >= 100:
            maturity_score += 3
        if back_speed > 0:
            maturity_score += 4
        if selected_routing:
            maturity_score += 4
        if ap_cnt > 0:
            maturity_score += 2
        if str(ngfw_vendor).strip().lower() not in {"", "нет", "no", "none"}:
            maturity_score += 5

    if server_active:
        server_total = phys_count + virt_count
        if phys_count > 0:
            maturity_score += 4
        if virt_count > 0:
            maturity_score += 5
        if server_total >= 5:
            maturity_score += 3
        if selected_virt_sys and virt_count > 0:
            maturity_score += 4
        if str(backup_vendor).strip().lower() not in {"", "нет", "no", "none"}:
            maturity_score += 4

    if storage_active:
        storage_disks = cnt_hdd + cnt_ssd
        if storage_disks > 0:
            maturity_score += 3
        if st_media_sel:
            maturity_score += 3
        if cnt_ssd > 0:
            maturity_score += 2
        if raid_selected:
            maturity_score += 4

    if systems_active:
        maturity_score += 6

    if web_active:
        maturity_score += 4

    if dev_active:
        if dev_count > 0:
            maturity_score += 4
        if dev_count >= 10:
            maturity_score += 3
        if sel_langs:
            maturity_score += 3
        if cicd_active:
            maturity_score += 3

    return min(100, maturity_score)


def build_context(results, client_info):

    users = int(results.get("_user_count", 0))

    industry = str(
        client_info.get(
            "Сфера деятельности",
            ""
        )
    )

    context = {}

    # ==================================
    # COMPANY SIZE
    # ==================================

    context["users"] = users

    context["small_company"] = users < 50

    context["medium_company"] = (
        50 <= users < 250
    )

    context["large_company"] = (
        250 <= users < 1000
    )

    context["enterprise_company"] = (
        users >= 1000
    )

    # ==================================
    # INDUSTRY
    # ==================================

    context["industry"] = industry

    context["is_finance"] = (
        "Финтех" in industry
        or "Банк" in industry
    )

    context["is_gov"] = (
        "Госсектор" in industry
    )

    context["is_it_company"] = (
        "IT" in industry
        or "Разработка" in industry
    )

    context["is_retail"] = (
        "Ритейл" in industry
        or "E-commerce" in industry
    )

    context["is_manufacturing"] = (
        "Производство" in industry
    )

    # ==================================
    # Бизнес-системы
    # ==================================

    result_text = str(results)

    context["has_erp"] = (
        "ERP" in result_text
    )

    context["has_crm"] = (
        "CRM" in result_text
    )

    context["has_accounting"] = (
        "Учет" in result_text
        or "Бухгалтерия" in result_text
    )

    context["has_hrm"] = (
        "HRM" in result_text
    )

    context["has_document_flow"] = (
        "СЭД" in result_text
    )

    context["has_mail"] = (
        results.get(
            "1.5.1. Почтовая система",
            "Нет"
        ) != "Нет"
    )

    # ==================================
    # CRITICAL DATA
    # ==================================

    context["has_personal_data"] = any([
        context["has_hrm"],
        context["has_crm"],
        context["has_accounting"]
    ])

    context["has_critical_systems"] = any([
        context["has_erp"],
        context["has_crm"],
        context["has_accounting"]
    ])

    # ==================================
    # Разработка
    # ==================================

    context["has_development"] = (
        "4.1. Разработчики" in results
    )

    dev_count = int(
        results.get(
            "4.1. Разработчики",
            0
        )
    )

    context["developers"] = dev_count

    context["large_dev_team"] = (
        dev_count >= 10
    )

    # ==================================
    # WEB
    # ==================================

    context["has_public_web"] = (
        "3.1. Хостинг" in results
    )

    # ==================================
    # Инфраструктура
    # ==================================

    context["servers"] = int(
        results.get(
            "Серверы (вирт)",
            0
        )
    ) + int(
        results.get(
            "Серверы (физ)",
            0
        )
    )

    context["has_virtualization"] = (
        context["servers"] > 0
    )

    context["has_backup"] = (
        results.get(
            "Резервное копирование",
            "Нет"
        ) != "Нет"
    )

    # ==================================
    # Безопасность
    # ==================================

    context["has_ngfw"] = (
        results.get(
            "NGFW",
            "Нет"
        ) != "Нет"
    )

    context["has_siem"] = (
        results.get(
            "SIEM",
            "Нет"
        ) != "Нет"
    )

    context["has_edr"] = (
        results.get(
            "EDR",
            "Нет"
        ) != "Нет"
    )

    context["has_patch_management"] = (
        results.get(
            "Patch Management",
            "Нет"
        ) != "Нет"
    )

    context["has_mfa"] = (
        results.get(
            "MFA",
            "Нет"
        ) != "Нет"
    )

    return context


def infrastructure_profile(context):
    users = context.get("users", 0)
    servers = context.get("servers", 0)

    if context.get("enterprise_company"):
        return "Крупная распределенная инфраструктура", (
            f"{users} АРМ, {servers} серверов. Требуется формальная модель "
            "управления ИБ, мониторинга, резервирования и регулярной отчетности."
        )

    if context.get("large_company"):
        return "Крупная инфраструктура", (
            f"{users} АРМ, {servers} серверов. Важно стандартизировать процессы, "
            "сегментацию, управление доступом и контроль обновлений."
        )

    if context.get("medium_company"):
        return "Средняя инфраструктура", (
            f"{users} АРМ, {servers} серверов. Приоритет - базовая управляемость, "
            "защита рабочих мест, резервное копирование и контроль доступа."
        )

    return "Малая инфраструктура", (
        f"{users} АРМ, {servers} серверов. Рекомендации должны быть практичными: "
        "минимум сложных платформ, максимум быстрых мер с понятной поддержкой."
    )


def build_contextual_roadmap(results, context, domain_scores, risks):
    roadmap = []

    def add(phase, priority, domain, action, rationale, owner="ИТ/ИБ", effort="Средняя"):
        roadmap.append({
            "phase": phase,
            "priority": priority,
            "domain": domain,
            "action": action,
            "rationale": rationale,
            "owner": owner,
            "effort": effort,
        })

    is_small = context.get("small_company", False)
    is_medium = context.get("medium_company", False)
    is_large = context.get("large_company", False)
    is_enterprise = context.get("enterprise_company", False)
    has_servers = context.get("servers", 0) > 0
    has_public_web = context.get("has_public_web", False)
    has_development = context.get("has_development", False)
    has_critical_systems = context.get("has_critical_systems", False)

    if results.get("MFA") == "Нет":
        add(
            "0-30 дней",
            "P1",
            "Доступ",
            "Включить MFA для администраторов, VPN, почты и критичных систем.",
            "Быстро снижает риск компрометации учетных записей без тяжелого проекта.",
            effort="Низкая",
        )

    if results.get("Резервное копирование") == "Нет" and has_servers:
        add(
            "0-30 дней",
            "P1",
            "Устойчивость",
            "Организовать регулярное резервное копирование серверов и критичных данных.",
            "Для серверной инфраструктуры backup является базовым условием восстановления после сбоя или ransomware.",
            effort="Средняя",
        )
    elif results.get("Резервное копирование") != "Нет" and results.get("Immutable Backup") == "Нет":
        add(
            "31-60 дней",
            "P2",
            "Устойчивость",
            "Добавить immutable/offline-копии и проверить сценарий восстановления.",
            "Обычные резервные копии могут быть удалены или зашифрованы при атаке.",
            effort="Средняя",
        )

    legacy_arm = results.get("ОС АРМ (Windows XP/Vista/7/8)", 0)
    legacy_srv = results.get("ОС Сервера (Windows Server 2008/2012 R2)", 0)
    if legacy_arm or legacy_srv:
        add(
            "0-30 дней",
            "P1",
            "Инфраструктура",
            "Составить план миграции или изоляции устаревших ОС.",
            f"Обнаружены устаревшие ОС: АРМ {legacy_arm}, серверы {legacy_srv}.",
            effort="Высокая",
        )

    if results.get("Patch Management") == "Нет":
        phase = "31-60 дней" if is_small else "0-30 дней"
        add(
            phase,
            "P2",
            "Управляемость",
            "Ввести централизованный контроль обновлений для АРМ и серверов.",
            "Чем больше инфраструктура, тем выше риск ручного и несогласованного обновления.",
            effort="Средняя",
        )

    if results.get("EDR") == "Нет":
        action = (
            "Выбрать управляемый EDR/MDR-сервис для рабочих мест."
            if is_small or is_medium
            else "Внедрить EDR/XDR с централизованной политикой и реагированием."
        )
        add(
            "31-60 дней",
            "P2",
            "Endpoint",
            action,
            "Обычного антивируса недостаточно для расследования сложных атак и lateral movement.",
            effort="Средняя",
        )

    if results.get("SIEM") == "Нет":
        if is_enterprise:
            action = "Запустить проект SIEM/SOC с корреляцией событий и регламентами реагирования."
            phase = "61-90 дней"
            priority = "P2"
            effort = "Высокая"
        elif is_large or has_critical_systems:
            action = "Подключить MSSP/SOC или легковесный SIEM для критичных систем и средств защиты."
            phase = "61-90 дней"
            priority = "P3"
            effort = "Средняя"
        else:
            action = "Настроить сбор критичных журналов и назначить ответственного за разбор инцидентов."
            phase = "31-60 дней"
            priority = "P3"
            effort = "Низкая"

        add(
            phase,
            priority,
            "Мониторинг",
            action,
            "Мониторинг должен соответствовать масштабу: для малой инфраструктуры важнее управляемый минимум, чем тяжелая SIEM-платформа.",
            effort=effort,
        )

    if has_public_web and results.get("WAF") == "Нет":
        add(
            "31-60 дней",
            "P2",
            "Веб",
            "Оценить необходимость WAF/CDN-защиты для публичных веб-ресурсов.",
            "Публичные сервисы увеличивают поверхность атаки и требуют отдельного контроля.",
            effort="Средняя",
        )

    if has_development and results.get("SAST") == "Нет" and results.get("DAST") == "Нет":
        add(
            "61-90 дней",
            "P3",
            "Разработка",
            "Встроить базовые SAST/DAST-проверки в процесс разработки.",
            "Для собственной разработки риски приложений лучше выявлять до публикации.",
            owner="Разработка/ИБ",
            effort="Средняя",
        )

    weakest_domains = [
        domain
        for domain, score in sorted(domain_scores.items(), key=lambda item: item[1])
        if score < 50
    ][:2]
    for domain in weakest_domains:
        if not any(item["domain"] == domain for item in roadmap):
            add(
                "61-90 дней",
                "P3",
                domain,
                f"Разработать целевой план улучшения домена: {domain}.",
                "Домен имеет низкую оценку и требует отдельного набора мер.",
                effort="Средняя",
            )

    risk_titles = {item.get("risk") for item in risks[:3]}
    if risk_titles and not roadmap:
        add(
            "0-30 дней",
            "P1",
            "Приоритизация",
            "Разобрать ключевые риски и назначить ответственных за их устранение.",
            "Анкета выявила риски, которые требуют управленческого решения.",
            effort="Низкая",
        )

    if not roadmap:
        add(
            "31-60 дней",
            "P3",
            "Развитие",
            "Провести контрольную проверку политик, backup и доступа администраторов.",
            "Даже при хорошем базовом уровне полезна регулярная проверка устойчивости.",
            effort="Низкая",
        )

    priority_order = {"P1": 1, "P2": 2, "P3": 3}
    phase_order = {"0-30 дней": 1, "31-60 дней": 2, "61-90 дней": 3}
    return sorted(
        roadmap,
        key=lambda item: (
            phase_order.get(item["phase"], 99),
            priority_order.get(item["priority"], 99),
        )
    )[:10]


def pick_catalog_vendors(catalog, keywords, fallback):
    selected = []
    for vendor in catalog:
        vendor_lower = vendor.lower()
        if any(keyword.lower() in vendor_lower for keyword in keywords):
            selected.append(vendor)

    if not selected:
        selected = fallback

    return list(dict.fromkeys(selected))[:4]


def normalize_vendor_key(value):
    return re.sub(r"[^a-zа-я0-9]+", " ", str(value or "").lower()).strip()


def clean_vendor_display_name(value):
    vendor = str(value or "").strip()
    fixes = {
        "imperva": "Imperva",
        "huawei": "Huawei",
        "splunc": "Splunk",
        "mccafee": "McAfee",
    }
    return fixes.get(vendor.lower(), vendor)


DETAILED_VENDOR_MATRIX_FILE = "vendor_matrix_detailed.xlsx"


def load_detailed_vendor_names():
    try:
        if not os.path.exists(DETAILED_VENDOR_MATRIX_FILE):
            return []
        df = pd.read_excel(DETAILED_VENDOR_MATRIX_FILE)
        if df.empty or "Vendor" not in df.columns:
            return []
        values = []
        for value in df["Vendor"].dropna().tolist():
            vendor = clean_vendor_display_name(value)
            if vendor and vendor.lower() not in {"nan", "none"}:
                values.append(vendor)
        return list(dict.fromkeys(values))
    except Exception:
        return []


def load_detailed_solution_vendor_map():
    try:
        if not os.path.exists(DETAILED_VENDOR_MATRIX_FILE):
            return {}
        df = pd.read_excel(DETAILED_VENDOR_MATRIX_FILE)
        if df.empty or "Vendor" not in df.columns:
            return {}

        category_map = {}
        for _, row in df.iterrows():
            vendor = clean_vendor_display_name(row.get("Vendor"))
            if not vendor:
                continue
            for column in df.columns:
                if column == "Vendor":
                    continue
                marker = str(row.get(column) or "").strip()
                if marker == "+":
                    category_map.setdefault(str(column).strip(), []).append(vendor)

        return {
            category: list(dict.fromkeys(vendors))
            for category, vendors in category_map.items()
        }
    except Exception:
        return {}


def normalize_portfolio_header(value):
    return normalize_vendor_key(value)


def split_portfolio_list(value):
    if value is None:
        return []
    parts = re.split(r"[,;|\n]+", str(value))
    return [part.strip() for part in parts if part and part.strip()]


def load_verified_distributor_map():
    try:
        if not os.path.exists(DETAILED_VENDOR_MATRIX_FILE):
            return {}
        df = pd.read_excel(DETAILED_VENDOR_MATRIX_FILE)
        if df.empty or "Vendor" not in df.columns:
            return {}

        normalized_columns = {
            normalize_portfolio_header(column): column
            for column in df.columns
        }
        distributor_column = None
        status_column = None
        for key, column in normalized_columns.items():
            if key in {
                "distributor", "distributors", "distributor name",
                "дистрибьютор", "дистрибьюторы", "поставщик"
            }:
                distributor_column = column
            if key in {
                "distributor status", "status distributor",
                "статус дистрибьютора", "статус"
            }:
                status_column = column

        if not distributor_column or not status_column:
            return {}

        verified_statuses = {
            "verified", "approved", "trusted", "checked",
            "проверенный", "проверено", "подтвержденный", "подтверждено",
        }
        distributor_map = {}
        for _, row in df.iterrows():
            status = normalize_portfolio_header(row.get(status_column))
            if status not in verified_statuses:
                continue
            vendor = clean_vendor_display_name(row.get("Vendor"))
            distributors = split_portfolio_list(row.get(distributor_column))
            if vendor and distributors:
                distributor_map.setdefault(vendor, [])
                distributor_map[vendor].extend(distributors)

        return {
            vendor: list(dict.fromkeys(distributors))
            for vendor, distributors in distributor_map.items()
        }
    except Exception:
        return {}


def verified_distributors_for_vendors(vendors_text):
    distributor_map = load_verified_distributor_map()
    if not distributor_map:
        return "-"
    vendors = split_portfolio_list(vendors_text)
    values = []
    for vendor in vendors:
        normalized_vendor = normalize_vendor_key(vendor)
        for known_vendor, distributors in distributor_map.items():
            if normalize_vendor_key(known_vendor) == normalized_vendor:
                values.append(f"{known_vendor}: {', '.join(distributors)}")
    return "\n".join(list(dict.fromkeys(values))) or "-"


def load_solution_vendor_map():
    solution_aliases = {
        "DLP": ("dlp", "утеч", "защита данных"),
        "Шифрование и маскирование данных": ("шифр", "маскир"),
        "Системы контроля доступа и управления привилегиями": (
            "pam", "iam", "privileged", "привилег", "учетн"
        ),
        "Архивирование и резервное копирование": (
            "backup", "резерв", "архив", "восстановлен", "rto", "rpo"
        ),
        "Защита баз данных и аудит": ("database", "баз дан", "db audit"),
        "NGFW": ("ngfw", "firewall", "межсет", "периметр"),
        "Сетевое оборудование": ("сетевое оборудование", "vlan", "маршрут", "коммутатор", "switch", "wifi", "wi fi"),
        "IDS/IPS": ("ids", "ips"),
        "WAF": ("waf", "web application security", "веб прилож", "owasp"),
        "AntiDDos": ("antiddos", "anti ddos", "anti-ddos", "ddos"),
        "Защита облаков": ("cloud", "облак", "casb"),
        "Cyber Risk Management": ("vulnerability", "уязв", "cve", "scanner", "сканер"),
        "Мониторинг и логирование": ("siem", "soar", "soc", "логирование", "журнал"),
        "Защита почты": ("mail security", "почт", "email", "фишинг"),
        "CASB, разработка и защита контейнеров": (
            "casb", "container", "контейнер", "sast", "dast", "appsec", "разработ"
        ),
        "MDM": ("mdm",),
        "Antiransomeware/EDR": ("edr", "xdr", "mdr", "epp", "endpoint", "конечн"),
        "Мультифаторная аутификация": ("mfa", "2fa", "многофактор"),
        "ITSM/CMDB": ("itsm", "cmdb", "change management", "configuration management", "управление изменениями", "управление конфигурациями"),
        "Миграция и виртуализация": ("миграция ос", "виртуализация", "virtualization", "migration project"),
    }
    detailed_matrix = load_detailed_solution_vendor_map()
    if detailed_matrix:
        solutions = list(solution_aliases)
        solution_categories = {
            solutions[0]: ("DLP",),
            solutions[1]: ("Encryption",),
            solutions[2]: ("PAM", "IGA"),
            solutions[3]: ("Backup", "Archiving"),
            solutions[4]: ("DAM/DB Security", "DB Security", "Data Discovery", "Data Classification"),
            solutions[5]: ("NGFW",),
            solutions[6]: ("Network Equipment",),
            solutions[7]: ("IDS/IPS",),
            solutions[8]: ("WAF",),
            solutions[9]: ("Anti-DDoS",),
            solutions[10]: ("Cloud Security", "CASB", "CNAPP", "CWPP"),
            solutions[11]: ("VM", "ASM", "Cyber Risk"),
            solutions[12]: ("SIEM", "SOAR", "UEBA"),
            solutions[13]: ("Email Security", "Email", "Corporate Email"),
            solutions[14]: ("CASB", "CWPP", "CNAPP"),
            solutions[15]: ("MDM", "MDM/UEM"),
            solutions[16]: ("EDR", "XDR", "AV"),
            solutions[17]: ("MFA",),
            solutions[18]: (),
            solutions[19]: ("Servers", "Storage", "Operating Systems", "Virtualization"),
        }
        vendor_map = {}
        for solution, categories in solution_categories.items():
            vendors = []
            for category in categories:
                vendors.extend(detailed_matrix.get(category, []))
            vendor_map[solution] = list(dict.fromkeys(vendors))
        return vendor_map, solution_aliases

    solution_vendor_keywords = {
        "DLP": ("symantec", "forcepoint", "zecurion", "ibatyr", "гарда"),
        "Шифрование и маскирование данных": ("symantec", "thales", "imperva", "гарда"),
        "Системы контроля доступа и управления привилегиями": (
            "cyberark", "axidian", "netwrix", "fudo"
        ),
        "Архивирование и резервное копирование": ("veeam", "commvault", "veritas"),
        "Защита баз данных и аудит": ("imperva", "гарда"),
        "NGFW": ("check point", "palo alto", "fortinet", "cisco", "huawei"),
        "Сетевое оборудование": ("cisco", "huawei", "juniper", "h3c", "hp", "dell"),
        "IDS/IPS": ("trend micro", "forcepoint", "check point", "fortinet"),
        "WAF": ("imperva", "f5", "cloudflare", "radware", "a10", "check point"),
        "AntiDDos": ("f5", "radware", "check point", "barracuda"),
        "Защита облаков": ("check point", "palo alto", "fortinet", "crowdstrike", "trend micro"),
        "Cyber Risk Management": ("tenable", "qualys", "positive", "harmony"),
        "Мониторинг и логирование": ("positive", "касперского", "netwrix", "splunc", "r vision", "qazsiem", "mccafee"),
        "Защита почты": ("trend micro", "barracuda", "forcepoint", "fortinet"),
        "CASB, разработка и защита контейнеров": (
            "check point", "palo alto", "fortinet", "forcepoint", "qualys", "hcl", "symantec", "trend micro"
        ),
        "MDM": ("kaspersky", "bitdefender", "eset"),
        "Antiransomeware/EDR": (
            "trend micro", "kaspersky", "check point", "symantec", "bitdefender", "fortinet", "eset", "crowdstrike"
        ),
        "Мультифаторная аутификация": ("eset", "axidian", "thales"),
        "ITSM/CMDB": (),
        "Миграция и виртуализация": (),
    }

    catalog = [
        clean_vendor_display_name(vendor)
        for vendor in load_vendor_names()
        if str(vendor).strip()
    ]
    vendor_map = {}

    for solution, keywords in solution_vendor_keywords.items():
        matched = []
        for vendor in catalog:
            normalized_vendor = normalize_vendor_key(vendor)
            if any(normalize_vendor_key(keyword) in normalized_vendor for keyword in keywords):
                matched.append(vendor)
        vendor_map[solution] = list(dict.fromkeys(matched))

    return vendor_map, solution_aliases


def manufacturers_for_report_item(item):
    vendor_map, solution_aliases = load_solution_vendor_map()
    existing_values = item.get("vendors", [])
    if not isinstance(existing_values, list):
        existing_values = [existing_values] if existing_values else []

    solution_text = normalize_vendor_key(" ".join(str(value) for value in existing_values))
    fallback_text = normalize_vendor_key(" ".join([
        str(item.get("risk", "")),
        str(item.get("description", "")),
    ]))

    manufacturers = []
    matched_solution = False

    detailed_matrix = load_detailed_solution_vendor_map()
    for category, vendors in detailed_matrix.items():
        if normalize_vendor_key(category) in solution_text:
            manufacturers.extend(vendors)
            matched_solution = True

    for solution, aliases in solution_aliases.items():
        if any(normalize_vendor_key(alias) in solution_text for alias in aliases):
            manufacturers.extend(vendor_map.get(solution, []))
            matched_solution = True

    if not matched_solution:
        for solution, aliases in solution_aliases.items():
            if any(normalize_vendor_key(alias) in fallback_text for alias in aliases):
                manufacturers.extend(vendor_map.get(solution, []))

    all_known_vendors = []
    for vendors in vendor_map.values():
        all_known_vendors.extend(vendors)

    for value in existing_values:
        normalized_value = normalize_vendor_key(value)
        for known_vendor in all_known_vendors:
            if normalize_vendor_key(known_vendor) == normalized_value:
                manufacturers.append(known_vendor)

    return ", ".join(list(dict.fromkeys(manufacturers))[:8]) or "-"


def build_sales_opportunities(results, context, roadmap_items):
    catalog = load_vendor_names()
    opportunities = []

    def add(priority, problem, offer, trigger, vendors, next_step):
        opportunities.append({
            "priority": priority,
            "problem": problem,
            "offer": offer,
            "trigger": trigger,
            "vendors": ", ".join(vendors),
            "next_step": next_step,
            "source": "Базовые правила",
        })

    if results.get("MFA") == "Нет":
        add(
            "P1",
            "Нет MFA для критичных доступов",
            "Проект MFA/IAM для администраторов, почты, VPN и критичных систем",
            "В анкете не указана многофакторная аутентификация.",
            pick_catalog_vendors(catalog, ["axidian", "cisco", "forti"], ["Cisco Duo", "FortiAuthenticator", "Axidian"]),
            "Уточнить текущий IdP/почтовую платформу и предложить пилот MFA без привязки к PAM-проекту.",
        )

    if results.get("EDR") == "Нет":
        add(
            "P1",
            "Есть EPP, но нет EDR/XDR",
            "Endpoint Detection & Response / MDR как развитие текущей защиты рабочих мест",
            "EPP закрывает базовую защиту, но не дает полноценного расследования атак.",
            pick_catalog_vendors(catalog, ["crowdstrike", "kaspersky", "trend", "bitdefender", "eset"], ["CrowdStrike", "Kaspersky", "Trend Micro"]),
            "Показать разницу EPP vs EDR на примерах ransomware/lateral movement.",
        )

    if results.get("Patch Management") == "Нет":
        add(
            "P1",
            "Нет централизованного Patch Management",
            "Проект управления обновлениями и уязвимостями",
            "Ручное обновление не масштабируется на текущий парк АРМ и серверов.",
            pick_catalog_vendors(catalog, ["tenable", "qualys", "positive", "kaspersky"], ["Tenable", "Qualys", "Positive Technologies"]),
            "Предложить инвентаризацию и отчет по критичным CVE за 1-2 недели.",
        )

    if results.get("SIEM") == "Нет" and not context.get("small_company"):
        add(
            "P2",
            "Нет централизованного мониторинга событий ИБ",
            "SIEM/SOC или MSSP-мониторинг для критичных систем",
            "Масштаб инфраструктуры и бизнес-системы требуют управляемого мониторинга.",
            pick_catalog_vendors(catalog, ["r-vision", "qazsiem", "r‑vision"], ["R-Vision SIEM", "QazSIEM"]),
            "Обсудить минимальный scope: NGFW, серверы, AD/учетки, EPP, backup.",
        )

    if results.get("WAF") == "Нет" and context.get("has_public_web"):
        add(
            "P2",
            "Публичные веб-ресурсы без WAF",
            "WAF/CDN/DDoS-защита для интернет-магазина и личного кабинета",
            "Публичная поверхность увеличивает риск атак на приложения и доступность.",
            pick_catalog_vendors(catalog, ["imperva", "f5", "cloudflare", "radware", "a10", "barracuda"], ["Imperva", "F5", "Cloudflare"]),
            "Предложить экспресс-аудит web-периметра и пилот WAF/CDN.",
        )

    if results.get("Резервное копирование") == "Нет":
        add(
            "P1",
            "Не указан backup-контур",
            "Резервное копирование и восстановление критичных сервисов",
            "Без backup бизнес не сможет гарантированно восстановиться после сбоя/ransomware.",
            pick_catalog_vendors(catalog, ["veeam", "commvault", "veritas"], ["Veeam", "Commvault", "Veritas"]),
            "Запросить RPO/RTO и предложить дизайн backup-политик.",
        )
    elif results.get("Immutable Backup") == "Нет":
        add(
            "P2",
            "Backup есть, но не указана защита копий от ransomware",
            "Immutable/offline backup и тест восстановления",
            "Обычные backup-копии могут быть удалены или зашифрованы атакующим.",
            pick_catalog_vendors(catalog, ["veeam", "commvault", "veritas"], ["Veeam", "Commvault", "Veritas"]),
            "Продать assessment текущего backup и сценарий контрольного восстановления.",
        )

    if context.get("servers", 0) >= 10 and results.get("DR") == "Нет":
        add(
            "P2",
            "Не описан DR-план для критичных ИТ-сервисов",
            "DR assessment, RTO/RPO-дизайн и тест восстановления",
            "Есть серверы, виртуализация и бизнес-системы, но не указан формализованный DR-сценарий.",
            pick_catalog_vendors(catalog, ["veeam", "commvault", "rubrik", "veritas"], ["Veeam", "Commvault", "Rubrik"]),
            "Запросить список критичных сервисов и предложить воркшоп по RTO/RPO с контрольным восстановлением.",
        )

    if context.get("servers", 0) >= 10 and results.get("Мониторинг") == "Нет":
        add(
            "P2",
            "Нет единого эксплуатационного мониторинга ИТ",
            "Мониторинг серверов, сети, СХД, виртуализации и бизнес-сервисов",
            "Инциденты по емкости, доступности и производительности могут обнаруживаться после влияния на пользователей.",
            pick_catalog_vendors(catalog, ["zabbix", "prtg", "manageengine"], ["Zabbix", "PRTG", "ManageEngine"]),
            "Предложить быстрый health-check мониторинга и пилот с 10-15 критичными объектами.",
        )

    if results.get("Виртуализация") != "Нет" and results.get("Кластеризация") == "Нет":
        add(
            "P3",
            "Нужно проверить отказоустойчивость виртуализации",
            "Аудит HA/DRS, резервов CPU/RAM/storage и размещения критичных VM",
            "Виртуальные серверы обслуживают бизнес-системы, но отказ хоста может затронуть сразу несколько сервисов.",
            pick_catalog_vendors(catalog, ["vmware", "veeam", "zabbix"], ["VMware", "Veeam", "Zabbix"]),
            "Предложить инфраструктурный аудит виртуализации и карту рисков по критичным VM.",
        )

    if results.get("СХД") != "Нет" and results.get("Мониторинг СХД") == "Нет":
        add(
            "P3",
            "Не описан capacity management СХД",
            "Контроль емкости, производительности и snapshot-политик СХД",
            "Рост данных или деградация RAID/latency может повлиять на ERP, CRM и файловые сервисы.",
            pick_catalog_vendors(catalog, ["zabbix", "prtg", "manageengine"], ["Zabbix", "PRTG", "ManageEngine"]),
            "Запросить модель СХД, текущую утилизацию и предложить capacity/performance assessment.",
        )

    if results.get("PAM") == "Нет" and (context.get("servers", 0) >= 10 or context.get("has_critical_systems")):
        add(
            "P2",
            "Нет PAM для привилегированных учетных записей",
            "Privileged Access Management",
            "Есть серверы/критичные системы, но не указан контроль администраторского доступа.",
            pick_catalog_vendors(catalog, ["cyberark", "fudo", "netwrix"], ["CyberArk", "Fudo", "Netwrix"]),
            "Предложить обследование админ-доступов и пилот vault/session recording.",
        )

    if not opportunities and roadmap_items:
        first = roadmap_items[0]
        add(
            first["priority"],
            "Требуется приоритизация улучшений",
            "Экспертный воркшоп и технический пресейл по roadmap",
            first["rationale"],
            catalog[:4] or ["Khalil Trade"],
            "Назначить встречу с ИТ/ИБ и пройти roadmap по шагам.",
        )

    priority_order = {"P1": 1, "P2": 2, "P3": 3}
    return sorted(opportunities, key=lambda item: priority_order.get(item["priority"], 99))[:10]


def build_sales_conversation_pack(c_info, results, context, roadmap_items, opportunities):
    company = c_info.get("Наименование компании", "клиент")
    users = context.get("users", 0)
    servers = context.get("servers", 0)
    profile_title, profile_text = infrastructure_profile(context)

    pains = []

    def add_pain(priority, pain, evidence, commercial_angle, discovery_question):
        pains.append({
            "priority": priority,
            "pain": pain,
            "evidence": evidence,
            "commercial_angle": commercial_angle,
            "discovery_question": discovery_question,
        })

    if results.get("MFA") == "Нет":
        add_pain(
            "P1",
            "Компрометация учетных записей остается одним из самых быстрых сценариев инцидента.",
            "В анкете не указана MFA для критичных доступов.",
            "MFA/IAM-пилот для администраторов, VPN, почты и критичных систем.",
            "Где сейчас находятся самые критичные учетные записи: почта, VPN, AD, ERP/CRM, облака?"
        )

    if results.get("Резервное копирование") == "Нет":
        add_pain(
            "P1",
            "Нет подтвержденного сценария восстановления после сбоя или ransomware.",
            f"Серверов указано: {servers}; backup-контур не указан.",
            "Backup assessment, дизайн RPO/RTO и внедрение резервного копирования.",
            "Какие сервисы бизнес должен восстановить первыми и за какое время?"
        )
    elif results.get("Immutable Backup") == "Нет":
        add_pain(
            "P2",
            "Backup есть, но его устойчивость к ransomware не подтверждена.",
            "Не указаны immutable/offline-копии и регулярный тест восстановления.",
            "Immutable backup, контрольное восстановление, регламент RTO/RPO.",
            "Когда последний раз проводили тестовое восстановление и кто подписал результат?"
        )

    if results.get("Patch Management") == "Нет":
        add_pain(
            "P1",
            "Риск эксплуатации известных уязвимостей растет быстрее, чем команда успевает обновлять вручную.",
            f"Масштаб: {users} АРМ, {servers} серверов; централизованный patch management не указан.",
            "Инвентаризация, vulnerability/patch management, регулярный отчет по критичным CVE.",
            "Как сейчас принимается решение, какие обновления ставить срочно, а какие можно отложить?"
        )

    if results.get("EDR") == "Нет":
        add_pain(
            "P1",
            "Команда может не увидеть сложную атаку на рабочих местах до влияния на бизнес.",
            "EDR/XDR/MDR не указаны; EPP без расследования и реагирования закрывает только базовый уровень.",
            "EDR/MDR-пилот на критичных группах пользователей и серверов.",
            "Есть ли сейчас возможность понять цепочку атаки: пользователь, файл, процесс, сеть, сервер?"
        )

    if results.get("SIEM") == "Нет" and not context.get("small_company"):
        add_pain(
            "P2",
            "События ИБ разрознены, поэтому инциденты сложно обнаруживать и расследовать.",
            "SIEM/SOC не указан при наличии инфраструктуры, требующей централизованного мониторинга.",
            "MSSP/SOC или поэтапное подключение SIEM для минимального critical scope.",
            "Какие источники логов сейчас реально просматриваются ежедневно, а какие только после инцидента?"
        )

    if context.get("has_public_web") and results.get("WAF") == "Нет":
        add_pain(
            "P2",
            "Публичные веб-сервисы увеличивают поверхность атаки и риск простоя.",
            "Публичный web-контур есть, WAF/CDN-защита не указана.",
            "Экспресс-аудит web-периметра, WAF/CDN/DDoS-пилот.",
            "Какие публичные сервисы приносят выручку или обслуживают клиентов напрямую?"
        )

    if context.get("servers", 0) >= 10 and results.get("Мониторинг") == "Нет":
        add_pain(
            "P2",
            "Инфраструктурные сбои могут обнаруживаться после жалоб пользователей.",
            "Есть серверный контур, но эксплуатационный мониторинг не указан.",
            "Мониторинг серверов, сети, СХД, виртуализации и бизнес-сервисов.",
            "Какие показатели сейчас отслеживаются: доступность, диски, latency, backup jobs, сервисы?"
        )

    if not pains:
        add_pain(
            "P3",
            "Анкета не показывает явных критичных разрывов, но нужен экспертный разбор целевой модели.",
            profile_text,
            "Пресейл-воркшоп по roadmap и проверка приоритетов.",
            "Какая зона сейчас больше всего беспокоит бизнес: доступность, безопасность, производительность или compliance?"
        )

    call_script = [
        {
            "stage": "Открытие",
            "talk_track": (
                f"Мы посмотрели анкету {company}. По масштабу это {profile_title.lower()}: "
                f"{users} АРМ и {servers} серверов. Цель звонка - не продавать набор продуктов, "
                "а подтвердить 2-3 приоритета, которые быстрее всего снизят риск."
            )
        },
        {
            "stage": "Подтверждение боли",
            "talk_track": (
                f"Главная гипотеза: {pains[0]['pain']} Основание: {pains[0]['evidence']} "
                "Правильно ли мы поняли ситуацию, или внутри есть компенсирующие меры?"
            )
        },
        {
            "stage": "Переход к решению",
            "talk_track": (
                f"Логичный первый шаг - {pains[0]['commercial_angle']} "
                "Мы можем начать с короткого assessment/pilot, чтобы не запускать тяжелый проект без подтверждения эффекта."
            )
        },
        {
            "stage": "Закрытие на следующий шаг",
            "talk_track": (
                "Предлагаю 45-минутную встречу с ИТ/ИБ: подтверждаем scope, фиксируем текущие ограничения "
                "и готовим короткий план действий/КП по первому приоритету."
            )
        },
    ]

    questions = []
    seen_questions = set()

    def add_question(topic, question):
        normalized = question.strip().lower()
        if normalized in seen_questions:
            return
        questions.append((topic, question))
        seen_questions.add(normalized)

    for pain in pains[:4]:
        add_question(pain["priority"] + " / " + pain["commercial_angle"].split(",")[0][:28], pain["discovery_question"])

    add_question("Приоритет бизнеса", "Какой риск из отчета для вас самый болезненный: простой сервиса, утечка данных, компрометация учеток или ручная эксплуатация?")
    add_question("Текущий бюджет", "На что уже заложен бюджет: продление текущих продуктов, новый пилот, аудит/assessment или сервисная модель?")
    add_question("Критичные сервисы", "Какие 3 системы нужно защищать и восстанавливать первыми, если случится инцидент?")

    if results.get("Patch Management") == "Нет":
        add_question("Patch / CVE", "Есть ли сейчас отчет, какие критичные CVE открыты на АРМ и серверах дольше допустимого срока?")
    if results.get("EDR") == "Нет":
        add_question("Endpoint", "Если на АРМ сработает подозрительный процесс, сможете ли вы восстановить цепочку: пользователь, файл, процесс, сеть, сервер?")
    if results.get("SIEM") == "Нет" and not context.get("small_company"):
        add_question("Логи / SOC", "Какие источники логов вы готовы подключить первыми: NGFW, AD/учетки, серверы, EPP, backup или почту?")
    if results.get("MFA") == "Нет":
        add_question("MFA / IAM", "Где MFA можно включить быстрее всего без ломки процессов: VPN, почта, администраторы, облака или бизнес-системы?")
    if results.get("Резервное копирование") != "Нет":
        add_question("Backup", "Когда последний раз делали тестовое восстановление и какой результат можно показать руководству?")
    if context.get("has_public_web"):
        add_question("Web", "Какие публичные приложения критичны для выручки или клиентского сервиса, и кто владелец их доступности?")
    if context.get("has_development"):
        add_question("Разработка", "Где в процессе релиза можно поставить security gate: зависимости, SAST, DAST или ручной review?")

    questions = questions[:10]

    objections = [
        (
            "У нас уже есть антивирус/NGFW",
            "Согласиться и развести уровни: EPP/NGFW - базовая защита, а вопрос отчета про обнаружение, расследование, доступы, backup и управляемость."
        ),
        (
            "Сейчас нет бюджета",
            "Предложить assessment/pilot с ограниченным scope и показать, какие риски можно закрыть быстро без enterprise-проекта."
        ),
        (
            "Мы маленькая компания, SIEM нам не нужен",
            "Подтвердить: тяжелый SIEM не нужен как первый шаг. Предложить минимальный сбор критичных логов или управляемый сервис."
        ),
        (
            "Все работает, инцидентов не было",
            "Сместить разговор на проверяемость: тест восстановления, отчет по критичным CVE, MFA для админов, мониторинг событий."
        ),
    ]

    next_steps = []
    for item in opportunities[:5]:
        next_steps.append({
            "priority": item["priority"],
            "step": item["next_step"],
            "offer": item["offer"],
            "success_criteria": "Подтвержден scope, владелец со стороны клиента и понятный следующий артефакт: пилот, assessment или КП.",
        })
    if not next_steps and roadmap_items:
        for item in roadmap_items[:3]:
            next_steps.append({
                "priority": item["priority"],
                "step": item["action"],
                "offer": "Экспертный воркшоп по roadmap",
                "success_criteria": "Клиент подтвердил приоритет и согласовал следующий созвон с техническими владельцами.",
            })

    return {
        "pains": pains[:8],
        "call_script": call_script,
        "questions": questions,
        "objections": objections,
        "next_steps": next_steps,
    }


def build_expert_conclusion(results, context, final_score, domain_scores, roadmap_items):
    profile_title, _ = infrastructure_profile(context)
    users = context.get("users", 0)
    servers = context.get("servers", 0)
    weak_domains = [
        domain
        for domain, score in sorted(domain_scores.items(), key=lambda item: item[1])
        if score < 50
    ]
    p1_actions = [
        item["action"]
        for item in roadmap_items
        if item["priority"] == "P1"
    ][:3]

    conclusion = [
        (
            f"Инфраструктура классифицирована как: {profile_title.lower()} "
            f"({users} АРМ, {servers} серверов). Рекомендации сформированы с учетом масштаба: "
            "для текущего профиля приоритет отдается мерам, которые быстро повышают управляемость, "
            "восстановимость и контроль доступа без избыточного внедрения enterprise-платформ."
        )
    ]

    if final_score < 35:
        conclusion.append(
            "Текущий уровень зрелости ИБ оценивается как начальный: ключевые защитные процессы "
            "фрагментарны, а эффективность реагирования зависит от ручных действий специалистов."
        )
    elif final_score < 60:
        conclusion.append(
            "Текущий уровень зрелости ИБ оценивается как базовый/развивающийся: отдельные средства "
            "защиты уже используются, но отсутствует целостная операционная модель контроля, "
            "обновлений, мониторинга и реагирования."
        )
    else:
        conclusion.append(
            "Текущий уровень зрелости ИБ достаточен для базового контроля, однако дальнейшее развитие "
            "должно быть направлено на устойчивость процессов, измеримость и регулярную проверку защиты."
        )

    if weak_domains:
        conclusion.append(
            "Наиболее слабые домены: "
            + ", ".join(weak_domains[:3])
            + ". Именно они должны лечь в основу первоочередного плана улучшений."
        )

    if results.get("MFA") == "Нет":
        conclusion.append(
            "Отсутствие MFA повышает риск компрометации учетных записей. Для бизнеса это один из "
            "самых быстрых и экономически оправданных шагов по снижению вероятности инцидента."
        )

    if results.get("EPP") != "Нет" and results.get("EDR") == "Нет":
        conclusion.append(
            "Наличие EPP закрывает базовую антивирусную защиту, но не решает задачу расследования "
            "сложных атак. Для текущего масштаба разумнее рассматривать EDR/MDR как следующий этап, "
            "а не как замену существующему EPP."
        )

    if results.get("SIEM") == "Нет":
        if context.get("small_company"):
            conclusion.append(
                "Полноценный SIEM не является первоочередной инвестицией для малого масштаба. "
                "Практичнее начать со сбора критичных журналов, ответственного за разбор событий "
                "и понятного регламента реагирования."
            )
        else:
            conclusion.append(
                "Для заданного масштаба требуется развитие централизованного мониторинга: "
                "минимум - MSSP/SOC или легковесный SIEM для критичных систем, максимум - "
                "полноценная корреляция событий и регламенты реагирования."
            )

    if p1_actions:
        conclusion.append(
            "Первоочередные действия на 0-30 дней: " + "; ".join(p1_actions) + "."
        )

    return conclusion[:7]


def is_enabled(value):

    if value is None:
        return False

    value = str(value).strip().lower()

    if value in ["", "нет", "none", "false", "-", "n/a"]:
        return False

    return True

def calculate_domain_scores(results):

    domains = {
        "Сетевая безопасность": 0,
        "Защита конечных точек": 0,
        "Идентификация и доступ": 0,
        "Мониторинг и SOC": 0,
        "Резервное копирование": 0,
        "Инфраструктура": 0
    }

    # =========================
    # Сетевая безопасность
    # =========================

    if is_enabled(results.get("NGFW")):
        domains["Сетевая безопасность"] += 25

    if is_enabled(results.get("WAF")):
        domains["Сетевая безопасность"] += 15

    if is_enabled(results.get("Anti-DDoS")):
        domains["Сетевая безопасность"] += 15

    if is_enabled(results.get("VPN")):
        domains["Сетевая безопасность"] += 10

    if is_enabled(results.get("NAC")):
        domains["Сетевая безопасность"] += 20

    if is_enabled(results.get("Сегментация сети")):
        domains["Сетевая безопасность"] += 15

    # =========================
    # Защита конечных устройств
    # =========================

    if is_enabled(results.get("Антивирус")):
        domains["Защита конечных точек"] += 20

    if is_enabled(results.get("EDR")):
        domains["Защита конечных точек"] += 40

    if is_enabled(results.get("Patch Management")):
        domains["Защита конечных точек"] += 20

    if is_enabled(results.get("MDM")):
        domains["Защита конечных точек"] += 10

    if is_enabled(results.get("Device Control")):
        domains["Защита конечных точек"] += 10

    # =========================
    # IAM
    # =========================

    if is_enabled(results.get("MFA")):
        domains["Идентификация и доступ"] += 35

    if is_enabled(results.get("IDM")):
        domains["Идентификация и доступ"] += 25

    if is_enabled(results.get("PAM")):
        domains["Идентификация и доступ"] += 25

    if is_enabled(results.get("SSO")):
        domains["Идентификация и доступ"] += 15

    # =========================
    # MONITORING
    # =========================

    if is_enabled(results.get("SIEM")):
        domains["Мониторинг и SOC"] += 40

    if is_enabled(results.get("SOC")):
        domains["Мониторинг и SOC"] += 30

    if is_enabled(results.get("NAD")):
        domains["Мониторинг и SOC"] += 15

    if is_enabled(results.get("Threat Intelligence")):
        domains["Мониторинг и SOC"] += 15

    # =========================
    # BACKUP
    # =========================

    if is_enabled(results.get("Резервное копирование")):
        domains["Резервное копирование"] += 30

    if is_enabled(results.get("Immutable Backup")):
        domains["Резервное копирование"] += 30

    if is_enabled(results.get("DR")):
        domains["Резервное копирование"] += 20

    if is_enabled(results.get("Air-Gap Backup")):
        domains["Резервное копирование"] += 20

    # =========================
    # INFRA
    # =========================

    if is_enabled(results.get("Виртуализация")):
        domains["Инфраструктура"] += 25

    if is_enabled(results.get("СХД")):
        domains["Инфраструктура"] += 25

    if is_enabled(results.get("Мониторинг")):
        domains["Инфраструктура"] += 20

    if is_enabled(results.get("Резервный канал")):
        domains["Инфраструктура"] += 15

    if is_enabled(results.get("Кластеризация")):
        domains["Инфраструктура"] += 15

    # Ограничение 100%
    for k in domains:
        domains[k] = min(domains[k], 100)

    return domains


def risk_level_label(level):
    labels = {
        "CRITICAL": "Критический",
        "HIGH": "Высокий",
        "MEDIUM": "Средний",
        "LOW": "Низкий",
        "КРИТИЧЕСКИЙ": "Критический",
        "ВЫСОКИЙ": "Высокий",
        "СРЕДНИЙ": "Средний",
        "НИЗКИЙ": "Низкий",
    }
    return labels.get(str(level).upper(), str(level))


def risk_source_label(source):
    if str(source).lower() in {"ai", "ии", "gemini"}:
        return "ИИ"
    return "Базовые правила"


def risk_semantic_key(item):
    text = " ".join(
        str(item.get(field, ""))
        for field in ("risk", "description", "impact", "recommendation")
    ).lower()

    buckets = [
        ("mfa", ("mfa", "многофактор", "2fa", "двухфактор")),
        ("legacy_os", ("legacy", "устаревш", "windows xp", "windows vista", "windows 7", "windows 8", "2008", "2012 r2")),
        ("siem_soc", ("siem", "soc", "мониторинг событий", "централизованный мониторинг")),
        ("patch", ("patch", "обновлен", "cve", "уязвим")),
        ("endpoint_detection", ("edr", "xdr", "endpoint", "рабочих мест", "lateral movement")),
        ("backup", ("backup", "резерв", "immutable", "ransomware")),
        ("web_waf", ("waf", "web", "веб", "owasp", "публичн")),
        ("pam", ("pam", "привилегирован", "администраторск")),
        ("dlp", ("dlp", "утеч", "конфиденциальн")),
        ("mail", ("mail", "почт", "фишинг")),
        ("segmentation", ("сегментац", "vlan", "lateral")),
        ("it_monitoring", ("эксплуатационный мониторинг", "доступности", "производительности", "capacity")),
        ("virtualization", ("виртуализац", "гипервизор", "vm", "хост")),
        ("storage", ("схд", "storage", "raid", "snapshot", "iops")),
        ("dr", ("dr", "аварийн", "rto", "rpo", "восстановлен")),
        ("business_systems", ("erp", "crm", "бизнес-систем")),
    ]

    for key, markers in buckets:
        if any(marker in text for marker in markers):
            return key

    title = str(item.get("risk", "")).strip().lower()
    return re.sub(r"\s+", " ", title)


def professionalize_risk_item(item, results, context):
    source = item.get("_source", "Базовые правила")
    key = risk_semantic_key(item)
    users = context.get("users", results.get("_user_count", 0))
    servers = context.get("servers", 0)
    legacy_arm = results.get("ОС АРМ (Windows XP/Vista/7/8)", 0)
    legacy_srv = results.get("ОС Сервера (Windows Server 2008/2012 R2)", 0)

    if risk_source_label(source) == "ИИ":
        normalized = dict(item)
        normalized["_source"] = source
        normalized.setdefault("level", "MEDIUM")
        normalized.setdefault("description", normalized.get("risk", "Риск требует дополнительного уточнения."))
        normalized.setdefault("impact", "Риск может повлиять на устойчивость ИТ/ИБ процессов.")
        normalized.setdefault("recommendation", "Уточнить текущий процесс, назначить ответственного и определить измеримый план улучшений.")
        if not normalized.get("vendors"):
            normalized["vendors"] = []
        if not normalized.get("regulators"):
            normalized["regulators"] = ["ISO 27001"]
        return normalized

    profiles = {
        "legacy_os": {
            "level": "CRITICAL",
            "risk": "Устаревшие операционные системы требуют миграции или изоляции",
            "description": f"В анкете указаны устаревшие ОС: АРМ - {legacy_arm}, серверы - {legacy_srv}. Такие системы не получают полноценные исправления безопасности и не должны рассматриваться как обычные рабочие места.",
            "impact": "Уязвимости устаревших ОС повышают риск компрометации, невозможности установки актуальных агентов защиты и остановки связанных бизнес-процессов.",
            "recommendation": "Составить реестр устаревших ОС; временно изолировать их отдельным сетевым сегментом и ограничить доступ; подготовить план миграции на поддерживаемые версии ОС с контрольной датой вывода из эксплуатации.",
            "vendors": ["Microsoft", "Red Hat", "VMware"],
            "regulators": ["CIS Controls", "ISO 27001", "NIST CSF"],
        },
        "endpoint_detection": {
            "level": "HIGH",
            "risk": "Защита рабочих мест не обеспечивает полноценное обнаружение и реагирование",
            "description": f"В инфраструктуре указано {users} АРМ, при этом EDR/XDR/MDR не описаны. Базовый EPP снижает риск массового malware, но не закрывает расследование сложных атак и lateral movement.",
            "impact": "Инцидент на рабочей станции может дольше оставаться незамеченным, распространяться на учетные записи и серверы, а расследование будет зависеть от ручного сбора данных.",
            "recommendation": "Проверить фактическое покрытие EPP и актуальность агентов; выбрать пилот EDR/MDR для критичных групп пользователей и серверов; внедрить регламент реагирования с метриками MTTD/MTTR и контрольной проверкой сценариев ransomware.",
            "vendors": ["CrowdStrike", "SentinelOne", "Trend Micro"],
            "regulators": ["CIS Controls", "NIST CSF", "MITRE ATT&CK"],
        },
        "mfa": {
            "level": "HIGH",
            "risk": "Критичные доступы не защищены многофакторной аутентификацией",
            "description": "В анкете не указана MFA для администраторов, почты, VPN/удаленного доступа и критичных бизнес-систем.",
            "impact": "Компрометация одного пароля может дать атакующему доступ к корпоративной почте, административным консолям или бизнес-приложениям.",
            "recommendation": "Включить MFA сначала для администраторов, почты и удаленного доступа; затем распространить на ERP/CRM и критичные облачные сервисы; контролировать исключения и ежемесячно выгружать отчет по пользователям без MFA.",
            "vendors": ["Cisco Duo", "FortiAuthenticator", "Axidian"],
            "regulators": ["ISO 27001", "CIS Controls", "NIST CSF"],
        },
        "siem_soc": {
            "level": "HIGH" if context.get("enterprise_company") else "MEDIUM",
            "risk": "Мониторинг событий ИБ не покрывает критичные источники",
            "description": "Не описан централизованный сбор и анализ событий от NGFW, серверов, учетных записей, endpoint-защиты, почты и публичных сервисов.",
            "impact": "Атаки и нарушения политик могут обнаруживаться поздно, а расследование будет затруднено из-за разрозненных журналов.",
            "recommendation": "Начать с минимального SOC/MSSP-scope: NGFW, AD/учетные записи, серверы, EPP/EDR, backup и почта; настроить 8-12 базовых сценариев корреляции; ежемесячно разбирать инциденты и качество источников логов.",
            "vendors": ["IBM QRadar", "Splunk", "R-Vision SIEM"],
            "regulators": ["ISO 27001", "NIST CSF", "CIS Controls"],
        },
        "web_waf": {
            "level": "HIGH",
            "risk": "Публичные веб-сервисы требуют специализированной защиты приложений",
            "description": "В анкете указаны интернет-магазин и личный кабинет, но WAF-защита и прикладные правила контроля атак не описаны.",
            "impact": "Публичные приложения остаются под риском OWASP-атак, бот-активности, утечки данных и нарушения доступности клиентских сервисов.",
            "recommendation": "Провести экспресс-оценку web-периметра; включить WAF/CDN для публичных ресурсов с профилем под приложение; настроить регулярный разбор блокировок и связать WAF-события с процессом реагирования.",
            "vendors": ["Imperva", "F5", "Cloudflare"],
            "regulators": ["OWASP ASVS", "PCI DSS", "ISO 27001"],
        },
        "pam": {
            "level": "HIGH",
            "risk": "Привилегированные учетные записи не выделены в отдельный контур контроля",
            "description": f"В инфраструктуре указано {servers} серверов и критичные бизнес-системы, но PAM или сопоставимый контроль административных доступов не описан.",
            "impact": "Компрометация одной администраторской учетной записи может привести к изменению конфигураций, отключению защитных средств или доступу к критичным данным.",
            "recommendation": "Провести инвентаризацию привилегированных учетных записей; внедрить vault и контроль сессий для критичных систем; настроить регулярный пересмотр прав и запрет постоянных локальных админов.",
            "vendors": ["CyberArk", "Fudo", "Netwrix"],
            "regulators": ["ISO 27001", "CIS Controls", "NIST CSF"],
        },
        "mail": {
            "level": "MEDIUM",
            "risk": "Почтовый контур требует отдельной anti-phishing защиты",
            "description": "В анкете указана корпоративная почта, но специализированная защита от фишинга, вредоносных вложений и подмены отправителя не описана.",
            "impact": "Фишинговая атака может стать первичной точкой компрометации учетных записей, рабочих мест и облачных сервисов.",
            "recommendation": "Проверить SPF/DKIM/DMARC и текущие политики фильтрации; внедрить mail security для вложений, URL и impersonation; проводить регулярные фишинг-симуляции и разбор результатов.",
            "vendors": ["Trend Micro", "Forcepoint", "Barracuda"],
            "regulators": ["CIS Controls", "NIST CSF"],
        },
        "patch": {
            "level": "HIGH" if users > 100 else "MEDIUM",
            "risk": "Управление обновлениями не выглядит централизованным",
            "description": f"Для парка {users} АРМ и {servers} серверов не описан управляемый процесс patch management и контроль устранения критичных CVE.",
            "impact": "Известные уязвимости могут оставаться открытыми дольше допустимого окна, повышая риск компрометации через типовые exploit-сценарии.",
            "recommendation": "Ввести ежемесячный цикл обновлений; отдельно контролировать критичные CVE с SLA; формировать отчет по покрытию АРМ, серверов и исключений после каждого окна обновлений.",
            "vendors": ["ManageEngine", "Ivanti", "Tanium"],
            "regulators": ["CIS Controls", "NIST CSF", "ISO 27001"],
        },
        "dr": {
            "level": "HIGH",
            "risk": "Не описан план аварийного восстановления критичных ИТ-сервисов",
            "description": "В анкете есть серверы, виртуализация, backup и бизнес-системы, но не описаны RTO/RPO, DR-сценарии и регулярные тесты восстановления.",
            "impact": "При сбое площадки, СХД или критичных виртуальных машин восстановление может занять дольше допустимого для бизнеса.",
            "recommendation": "Определить перечень критичных сервисов и целевые RTO/RPO; провести тест восстановления ERP/CRM/почты; оформить DR-runbook и назначить периодичность контрольных учений.",
            "vendors": ["Veeam", "Commvault", "Rubrik"],
            "regulators": ["ISO 22301", "ISO 27001", "NIST CSF"],
        },
        "it_monitoring": {
            "level": "MEDIUM",
            "risk": "Эксплуатационный мониторинг ИТ-инфраструктуры требует формализации",
            "description": "В анкете есть сеть, серверы, СХД и бизнес-системы, но не описан единый мониторинг доступности, производительности и емкости.",
            "impact": "Инциденты по дискам, каналам связи, виртуальным машинам или бизнес-сервисам могут обнаруживаться постфактум по обращениям пользователей.",
            "recommendation": "Настроить мониторинг серверов, сетевых устройств, СХД, виртуализации и ключевых приложений; определить пороги и ответственных; ежемесячно анализировать тренды емкости и доступности.",
            "vendors": ["Zabbix", "PRTG", "ManageEngine"],
            "regulators": ["ITIL", "ISO 20000", "ISO 27001"],
        },
        "virtualization": {
            "level": "MEDIUM",
            "risk": "Отказоустойчивость виртуализации требует отдельной проверки",
            "description": "Используются виртуальные серверы, но в анкете не описаны HA/DRS-настройки, резервы ресурсов и правила размещения критичных VM.",
            "impact": "Отказ одного хоста или нехватка ресурсов может затронуть сразу несколько бизнес-сервисов.",
            "recommendation": "Проверить HA-настройки и резервы CPU/RAM/storage; разделить критичные VM по хостам; оформить регламент обслуживания гипервизоров без простоя сервисов.",
            "vendors": ["VMware", "Veeam", "Zabbix"],
            "regulators": ["ITIL", "ISO 20000"],
        },
        "storage": {
            "level": "MEDIUM",
            "risk": "Контроль емкости и производительности СХД не описан как процесс",
            "description": "В анкете указана гибридная СХД, но не описаны пороги заполнения, latency/IOPS, snapshot-политики и план расширения емкости.",
            "impact": "Рост данных или деградация производительности может повлиять на ERP, CRM, файловые сервисы и виртуальные машины.",
            "recommendation": "Ввести capacity management для СХД; контролировать заполнение, latency и состояние RAID-групп; ежеквартально обновлять план расширения емкости и snapshot-политики.",
            "vendors": ["Zabbix", "PRTG", "ManageEngine"],
            "regulators": ["ITIL", "ISO 20000"],
        },
    }

    profile = profiles.get(key)
    if profile:
        normalized = {**item, **profile}
        normalized["_source"] = source
        return normalized

    normalized = dict(item)
    normalized["_source"] = source
    normalized.setdefault("level", "MEDIUM")
    normalized.setdefault("description", normalized.get("risk", "Риск требует дополнительного уточнения."))
    normalized.setdefault("impact", "Риск может повлиять на устойчивость ИТ/ИБ процессов.")
    normalized.setdefault("recommendation", "Уточнить текущий процесс, назначить ответственного и определить измеримый план улучшений.")
    if not normalized.get("vendors"):
        normalized["vendors"] = []
    if not normalized.get("regulators"):
        normalized["regulators"] = ["ISO 27001"]
    return normalized


def build_report_risk_set(c_info, results, context):
    rule_risks = generate_rule_based_risks(
        results,
        context
    )
    rule_risks = [
        {**item, "_source": "Базовые правила"}
        for item in rule_risks
        if isinstance(item, dict)
    ]

    ai_risks = ai_generate_risks_and_recs(
        c_info,
        results
    )
    ai_used = isinstance(ai_risks, list) and any(isinstance(item, dict) for item in ai_risks)
    st.session_state.ai_used_in_last_report = ai_used

    combined_risks = []
    if ai_used:
        combined_risks.extend([
            {**item, "_source": "ИИ"}
            for item in ai_risks
            if isinstance(item, dict)
        ])
    combined_risks.extend(rule_risks)

    priority_order = {"CRITICAL": 1, "HIGH": 2, "MEDIUM": 3, "LOW": 4}
    unique_risks = []
    seen_risks = set()
    skipped_conflicting_risks = []
    for item in combined_risks:
        conflict = risk_conflicts_with_answers(item, results)
        if conflict:
            skipped_conflicting_risks.append(
                f"{risk_semantic_key(item)}: {conflict}"
            )
            continue
        item = professionalize_risk_item(item, results, context)
        semantic_key = risk_semantic_key(item)
        if not semantic_key or semantic_key in seen_risks:
            continue
        unique_risks.append(item)
        seen_risks.add(semantic_key)

    st.session_state.last_report_skipped_conflicts = skipped_conflicting_risks

    sorted_risks = sorted(
        unique_risks,
        key=lambda item: priority_order.get(str(item.get("level", "")).upper(), 99)
    )
    required_it_risks = [
        item
        for item in sorted_risks
        if item.get("_ai_area") == "ИТ"
    ][:4]

    report_risks = []
    report_keys = set()
    for item in [*required_it_risks, *sorted_risks]:
        key = risk_semantic_key(item)
        if not key or key in report_keys:
            continue
        report_risks.append(item)
        report_keys.add(key)
        if len(report_risks) >= 12:
            break

    st.session_state.last_report_risk_sources = [
        {
            "level": risk_level_label(item.get("level", "MEDIUM")),
            "risk": item.get("risk", "Риск"),
            "recommendation": item.get("recommendation", "-"),
            "source": risk_source_label(item.get("_source")),
        }
        for item in report_risks
    ]

    return report_risks, ai_used


def build_audit_observations(results, context, domain_scores, report_risks):
    users = context.get("users", 0)
    servers = context.get("servers", 0)
    weak_domains = [
        domain
        for domain, score in sorted(domain_scores.items(), key=lambda item: item[1])
        if score < 50
    ][:4]

    observations = [
        (
            "Масштаб и управляемость",
            f"Инфраструктура уже вышла за рамки малого контура: {users} АРМ и {servers} серверов требуют формализованных процессов обновлений, мониторинга, резервного копирования и контроля изменений."
        )
    ]

    if weak_domains:
        observations.append((
            "Слабые домены защиты",
            "Наиболее слабые зоны по анкете: " + ", ".join(weak_domains) + ". Их нужно закрывать как программу улучшений, а не отдельными разрозненными закупками."
        ))

    if results.get("Резервное копирование") != "Нет":
        observations.append((
            "Устойчивость и восстановление",
            "Backup-контур заявлен, но для управленческого уровня важно проверить RTO/RPO, immutable/offline-копии и факт регулярного тестового восстановления."
        ))

    if context.get("has_public_web"):
        observations.append((
            "Публичная поверхность",
            "Интернет-магазин и личный кабинет формируют внешний периметр риска: защиту web-приложений, журналирование и реагирование нужно выделить в отдельный поток работ."
        ))

    if context.get("has_development"):
        observations.append((
            "Разработка и изменения",
            "Наличие внутренней разработки означает, что часть рисков нужно закрывать до публикации: через SAST/DAST, контроль релизов и требования безопасности к CI/CD."
        ))

    if report_risks:
        observations.append((
            "Фокус первых действий",
            "Первые управленческие решения должны закрыть доступы, устаревшие ОС, endpoint detection, patch management и web-периметр."
        ))

    return observations[:6]


def build_target_operating_model(results, context):
    users = context.get("users", 0)
    servers = context.get("servers", 0)
    access_text = (
        "Контроль покрытия MFA для администраторов, почты, удаленного доступа и критичных систем; регулярный пересмотр исключений и привилегий."
        if is_enabled(results.get("MFA"))
        else "MFA для администраторов, почты, удаленного доступа и критичных систем; регулярный пересмотр исключений и привилегий."
    )
    endpoint_text = (
        f"Единое покрытие EPP/EDR для {users} АРМ с контролем статуса агентов, инцидентов и реакции на ransomware-сценарии."
        if not any(is_enabled(results.get(control)) for control in ("EDR", "XDR", "MDR"))
        else f"Контроль покрытия endpoint-защиты для {users} АРМ: статус агентов, качество телеметрии, сценарии реагирования и метрики MTTD/MTTR."
    )
    return [
        ("Доступы", access_text),
        ("Рабочие места", endpoint_text),
        ("Инфраструктура", f"Мониторинг серверов, виртуализации, СХД и каналов связи для {servers} серверов с порогами, владельцами и отчетом по доступности."),
        ("Восстановление", "Проверенные RTO/RPO, immutable/offline backup и регулярный тест восстановления критичных бизнес-систем."),
        ("Публичные сервисы", "WAF/CDN, журналирование web-событий, разбор блокировок и связка с процессом реагирования."),
        ("Процессы", "Ежемесячный цикл patch management, управление изменениями, контроль исключений и регулярный управленческий отчет по рискам."),
    ]


def build_management_decisions(results, context):
    decisions = [
        "Утвердить владельцев направлений: доступы, рабочие места, инфраструктура, backup/DR, web-периметр.",
        "Провести пилот EDR/MDR на критичных группах пользователей и серверах, затем принять решение о масштабировании.",
        "Определить минимальный SOC/MSSP-scope: NGFW, AD/учетки, серверы, endpoint, backup, почта и web.",
        "Согласовать регулярный отчет для руководства: остаточные риски, выполненные меры, исключения, SLA закрытия критичных уязвимостей.",
    ]
    if is_enabled(results.get("MFA")):
        decisions.insert(
            1,
            "Запустить 30-дневный план: проверить покрытие MFA по критичным доступам, закрыть устаревшие ОС, формализовать patch management и проверить backup-восстановление."
        )
    else:
        decisions.insert(
            1,
            "Запустить 30-дневный план: включить MFA для критичных доступов, закрыть устаревшие ОС, формализовать patch management и проверить backup-восстановление."
        )
    if context.get("has_development"):
        decisions.append("Добавить требования AppSec в процесс релизов: SAST/DAST, проверка зависимостей и критерии допуска в продуктив.")
    return decisions[:6]


# --- Отчет ---
def make_expert_excel(c_info, results, final_score):
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "01 Заключение"
    # =========================
    # CORPORATE STYLES
    # =========================

    dark_blue_fill = PatternFill(start_color="1F3A5F", end_color="1F3A5F", fill_type="solid")
    light_blue_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    gray_fill = PatternFill(start_color="F3F6F9", end_color="F3F6F9", fill_type="solid")
    critical_fill = PatternFill(start_color="FDE9E7", end_color="FDE9E7", fill_type="solid")
    medium_fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")

    white_font = Font(color="FFFFFF", bold=True)
    section_font = Font(size=14, bold=True, color="1F1F1F")
    normal_font = Font(size=11)

    border = Border(
        left=Side(style='thin', color="D9D9D9"),
        right=Side(style='thin', color="D9D9D9"),
        top=Side(style='thin', color="D9D9D9"),
        bottom=Side(style='thin', color="D9D9D9")
    )
    maturity, maturity_icon = get_maturity_level(final_score)
    domain_scores = calculate_domain_scores(results)
    context = build_context(results, c_info)
    profile_title, profile_text = infrastructure_profile(context)
    report_risks, ai_used = build_report_risk_set(c_info, results, context)
    ai_narrative = st.session_state.get("ai_audit_narrative", {}) if ai_used else {}
    top_risks = generate_rule_based_risks(
        results,
        context
    )
    roadmap_items = ai_narrative.get("roadmap") or build_contextual_roadmap(
        results,
        context,
        domain_scores,
        top_risks
    )
    expert_conclusion = build_expert_conclusion(
        results,
        context,
        final_score,
        domain_scores,
        roadmap_items
    )

    # =========================
    # EXECUTIVE SUMMARY
    # =========================

    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ"
    ws['A1'].font = Font(bold=True, size=20, color="1F1F1F")

    ws.merge_cells('A3:D3')
    ws['A3'] = "ПАСПОРТ КОМПАНИИ И АУДИТА"
    ws['A3'].font = white_font
    ws['A3'].fill = dark_blue_fill
    ws['A3'].alignment = Alignment(horizontal='center')

    company_rows = [
        ("Компания", c_info.get('Наименование компании', '-'), "Дата аудита", datetime.now().strftime('%d.%m.%Y')),
        ("Сфера", c_info.get('Сфера деятельности', '-'), "Город", c_info.get('Город', '-')),
        ("Сайт", c_info.get('Сайт компании', '-'), "Контакт", c_info.get('ФИО контактного лица', '-')),
        ("Зрелость ИБ", f"{maturity_icon} {maturity} / {final_score}%", "Профиль инфраструктуры", profile_title),
        ("Масштаб", profile_text, "Формат", "Автоматический экспертный аудит"),
    ]

    for row_idx, row_values in enumerate(company_rows, start=4):
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if col_idx in [1, 3]:
                cell.font = Font(bold=True)
                cell.fill = light_blue_fill
            else:
                cell.fill = gray_fill

    # Executive Summary Block
    ws.merge_cells('A11:D11')
    ws['A11'] = "УПРАВЛЕНЧЕСКОЕ РЕЗЮМЕ"
    ws['A11'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A11'].fill = dark_blue_fill
    ws['A11'].alignment = Alignment(horizontal='center')

    summary_text = []
    if ai_narrative.get("executive_summary"):
        summary_text = [
            f"• {item}"
            for item in ai_narrative.get("executive_summary", [])[:7]
            if str(item).strip()
        ]
    else:
        summary_text.append(f"• Профиль инфраструктуры: {profile_text}")
        summary_text.append("• Общий вывод: текущий уровень зрелости требует не точечных закупок, а короткой программы стабилизации ИТ и ИБ с владельцами, сроками и метриками контроля.")

        if results.get("NGFW") != "Нет":
            summary_text.append("• Сильная сторона: используется NGFW, его нужно включить в единый контур мониторинга, журналирования и реагирования.")

        if results.get("MFA") == "Нет":
            summary_text.append("• Критичный разрыв: MFA не указана для критичных доступов, поэтому компрометация пароля остается одним из наиболее вероятных сценариев инцидента.")

        if results.get("SIEM") == "Нет":
            if context.get("small_company"):
                summary_text.append("• Для текущего масштаба приоритетнее сбор критичных журналов и регламент реакции, чем тяжелый SIEM-проект")
            else:
                summary_text.append("• Требуется централизованный мониторинг событий по минимальному scope: NGFW, учетные записи, серверы, endpoint, backup, почта и web.")

        if results.get("Резервное копирование") == "Нет":
            summary_text.append("• Не обнаружено централизованное резервное копирование")
        else:
            summary_text.append("• Backup заявлен, но управленчески важно подтвердить RTO/RPO, immutable/offline-копии и результаты тестового восстановления.")

        if results.get("_user_count", 0) > 100:
            summary_text.append("• Масштаб 100+ АРМ требует формализации patch management, endpoint response, мониторинга инфраструктуры и регулярного отчета по остаточным рискам.")

        if not summary_text:
            summary_text.append("• Базовые меры информационной безопасности реализованы")

    ws.merge_cells('A12:D19')
    ws['A12'] = "\n".join(summary_text)
    ws['A12'].alignment = Alignment(wrap_text=True, vertical='top')

    for row in range(12, 20):
        for col in range(1, 5):
            ws.cell(row=row, column=col).fill = gray_fill
            ws.cell(row=row, column=col).border = border

    ws['A12'].font = Font(size=12)

    current_row = 21

    # =========================
    # EXPERT CONCLUSION
    # =========================

    ws.merge_cells(f'A{current_row}:D{current_row}')
    conclusion_header = ws.cell(row=current_row, column=1, value="ЭКСПЕРТНОЕ ЗАКЛЮЧЕНИЕ")
    conclusion_header.font = white_font
    conclusion_header.fill = dark_blue_fill
    conclusion_header.alignment = Alignment(horizontal='center')
    current_row += 1

    conclusion_end_row = current_row + 8
    ws.merge_cells(start_row=current_row, start_column=1, end_row=conclusion_end_row, end_column=4)
    conclusion_cell = ws.cell(row=current_row, column=1, value="\n\n".join(expert_conclusion))
    conclusion_cell.alignment = Alignment(wrap_text=True, vertical='top')
    conclusion_cell.font = Font(size=11)

    for row in range(current_row, conclusion_end_row + 1):
        for col in range(1, 5):
            ws.cell(row=row, column=col).fill = gray_fill
            ws.cell(row=row, column=col).border = border

    observations = ai_narrative.get("audit_observations") or build_audit_observations(results, context, domain_scores, report_risks)
    target_model = build_target_operating_model(results, context)
    management_decisions = ai_narrative.get("management_decisions") or build_management_decisions(results, context)

    current_row = conclusion_end_row + 2

    # =========================
    # AUDITOR OBSERVATIONS
    # =========================
    ws.merge_cells(f'A{current_row}:D{current_row}')
    obs_header = ws.cell(row=current_row, column=1, value="КЛЮЧЕВЫЕ НАБЛЮДЕНИЯ АУДИТОРА")
    obs_header.font = white_font
    obs_header.fill = dark_blue_fill
    obs_header.alignment = Alignment(horizontal='center')
    current_row += 1

    for title, body in observations:
        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True)
        ws.cell(row=current_row, column=1).fill = light_blue_fill
        ws.cell(row=current_row, column=1).border = border
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
        body_cell = ws.cell(row=current_row, column=2, value=body)
        body_cell.alignment = Alignment(wrap_text=True, vertical='top')
        body_cell.border = border
        body_cell.fill = gray_fill
        for col in range(3, 5):
            ws.cell(row=current_row, column=col).border = border
            ws.cell(row=current_row, column=col).fill = gray_fill
        current_row += 1

    current_row += 1

    # =========================
    # TARGET OPERATING MODEL
    # =========================
    ws.merge_cells(f'A{current_row}:D{current_row}')
    tom_header = ws.cell(row=current_row, column=1, value="ЦЕЛЕВАЯ МОДЕЛЬ УЛУЧШЕНИЙ")
    tom_header.font = white_font
    tom_header.fill = dark_blue_fill
    tom_header.alignment = Alignment(horizontal='center')
    current_row += 1

    for title, body in target_model:
        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True)
        ws.cell(row=current_row, column=1).fill = light_blue_fill
        ws.cell(row=current_row, column=1).border = border
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
        body_cell = ws.cell(row=current_row, column=2, value=body)
        body_cell.alignment = Alignment(wrap_text=True, vertical='top')
        body_cell.border = border
        body_cell.fill = gray_fill
        for col in range(3, 5):
            ws.cell(row=current_row, column=col).border = border
            ws.cell(row=current_row, column=col).fill = gray_fill
        current_row += 1

    current_row += 1

    # =========================
    # MANAGEMENT DECISIONS
    # =========================
    ws.merge_cells(f'A{current_row}:D{current_row}')
    dec_header = ws.cell(row=current_row, column=1, value="ПЕРВЫЕ УПРАВЛЕНЧЕСКИЕ РЕШЕНИЯ")
    dec_header.font = white_font
    dec_header.fill = dark_blue_fill
    dec_header.alignment = Alignment(horizontal='center')
    current_row += 1

    for idx, decision in enumerate(management_decisions, start=1):
        ws.cell(row=current_row, column=1, value=idx).font = Font(bold=True)
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal='center')
        ws.cell(row=current_row, column=1).fill = light_blue_fill
        ws.cell(row=current_row, column=1).border = border
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
        decision_cell = ws.cell(row=current_row, column=2, value=decision)
        decision_cell.alignment = Alignment(wrap_text=True, vertical='top')
        decision_cell.fill = gray_fill
        decision_cell.border = border
        for col in range(3, 5):
            ws.cell(row=current_row, column=col).border = border
            ws.cell(row=current_row, column=col).fill = gray_fill
        current_row += 1

    current_row += 2

    # =========================
    # 7 / 30 / 90 DAY ACTION PLAN
    # =========================
    ws.merge_cells(f'A{current_row}:D{current_row}')
    plan_header = ws.cell(row=current_row, column=1, value="ПЛАН ДЕЙСТВИЙ: 7 / 30 / 90 ДНЕЙ")
    plan_header.font = white_font
    plan_header.fill = dark_blue_fill
    plan_header.alignment = Alignment(horizontal='center')
    current_row += 1

    action_plan = [
        (
            "Первые 7 дней",
            "Подтвердить владельцев рисков, критичные сервисы, RTO/RPO, администраторские доступы и текущий порядок обновлений.",
            "Появляется единый список приоритетов и ответственных, без запуска тяжелого проекта."
        ),
        (
            "До 30 дней",
            "; ".join(item["action"] for item in roadmap_items if item["phase"] == "0-30 дней") or
            "Закрыть быстрые меры: MFA для критичных доступов, проверка backup, инвентаризация активов и критичных уязвимостей.",
            "Снижаются наиболее вероятные сценарии инцидентов и появляется измеримый baseline."
        ),
        (
            "До 90 дней",
            "; ".join(item["action"] for item in roadmap_items if item["phase"] in {"31-60 дней", "61-90 дней"})[:420] or
            "Перейти от точечных мер к управляемой модели мониторинга, обновлений, восстановления и реагирования.",
            "Формируется программа улучшений с метриками контроля и понятным бюджетированием."
        ),
    ]

    for period, action, result in action_plan:
        ws.cell(row=current_row, column=1, value=period).font = Font(bold=True)
        ws.cell(row=current_row, column=1).fill = light_blue_fill
        ws.cell(row=current_row, column=1).border = border
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
        action_cell = ws.cell(row=current_row, column=2, value=action)
        action_cell.alignment = Alignment(wrap_text=True, vertical='top')
        action_cell.fill = gray_fill
        action_cell.border = border
        result_cell = ws.cell(row=current_row, column=4, value=result)
        result_cell.alignment = Alignment(wrap_text=True, vertical='top')
        result_cell.fill = gray_fill
        result_cell.border = border
        for col in range(3, 4):
            ws.cell(row=current_row, column=col).fill = gray_fill
            ws.cell(row=current_row, column=col).border = border
        current_row += 1

    current_row += 2

    # =========================
    # TOP RISKS OVERVIEW
    # =========================
    ws.merge_cells(f'A{current_row}:D{current_row}')

    risk_header = ws.cell(
        row=current_row,
        column=1,
        value="КЛЮЧЕВЫЕ РИСКИ ИТ И ИБ"
    )

    risk_header.font = white_font
    risk_header.fill = dark_blue_fill
    risk_header.alignment = Alignment(horizontal='center')

    current_row += 1

    headers = ["#", "Риск", "Критичность", "Влияние на бизнес"]

    for col_num, header in enumerate(headers, 1):

        cell = ws.cell(
            row=current_row,
            column=col_num,
            value=header
        )

        cell.font = white_font
        cell.fill = dark_blue_fill
        cell.border = border

    current_row += 1

    for idx, risk in enumerate(report_risks[:5], start=1):

        level = risk.get("level", "MEDIUM")
        impact = risk.get("impact", "-")

        ws.cell(row=current_row, column=1, value=idx).border = border
        ws.cell(row=current_row, column=2, value=risk.get("risk", "-")).border = border
        ws.cell(row=current_row, column=3, value=risk_level_label(level)).border = border
        ws.cell(row=current_row, column=4, value=impact).border = border

        if "CRITICAL" in str(level).upper():
            fill = critical_fill
        else:
            fill = medium_fill

        for c in range(1, 5):
            ws.cell(row=current_row, column=c).fill = fill

        current_row += 1

    current_row += 2

    # =========================
    # Оценка доменов безопасности
    # =========================

    ws.merge_cells(f'A{current_row}:D{current_row}')
    dom_cell = ws.cell(row=current_row, column=1, value="ОЦЕНКА ДОМЕНОВ БЕЗОПАСНОСТИ")
    dom_cell.font = white_font
    dom_cell.fill = dark_blue_fill
    dom_cell.alignment = Alignment(horizontal='center')

    current_row += 1

    headers = ["Домен безопасности", "Оценка", "Статус"]

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=col_num, value=header)
        cell.font = white_font
        cell.fill = dark_blue_fill
        cell.border = border

    current_row += 1

    for domain, score in domain_scores.items():

        if score >= 80:
            status = "🟢 Сильный"
            fill = light_blue_fill

        elif score >= 50:
            status = "🟠 Средний"
            fill = medium_fill

        else:
            status = "🔴 Слабый"
            fill = critical_fill

        ws.cell(row=current_row, column=1, value=domain).border = border
        ws.cell(row=current_row, column=2, value=f"{score}%").border = border
        ws.cell(row=current_row, column=3, value=status).border = border

        for c in range(1, 4):
            ws.cell(row=current_row, column=c).fill = fill

        current_row += 1

    current_row += 2

    curr_row = current_row
    ws.merge_cells(f'A{curr_row}:B{curr_row}')
    ws.cell(row=curr_row, column=1, value="ВЫЯВЛЕННЫЕ РИСКИ И РЕКОМЕНДАЦИИ").font = Font(bold=True, size=14)
    curr_row += 1

    ai_data = report_risks

    if ai_data:
        for item in ai_data:

            # Уровень и Название
            lvl = item.get('level', 'СРЕДНИЙ')

            ws.merge_cells(f'A{curr_row}:B{curr_row}')

            cell = ws.cell(
                row=curr_row,
                column=1,
                value=f"[{risk_level_label(lvl)}] {item.get('risk', 'Риск')}"
            )

            cell.font = Font(bold=True)

            if "КРИТ" in str(lvl).upper() or "CRITICAL" in str(lvl).upper():
                cell.fill = critical_fill
            else:
                cell.fill = medium_fill

            curr_row += 1

            # Описание, Влияние, Рекомендация
            fields = [
                ("Описание", item.get('description', '-')),
                ("Влияние", item.get('impact', '-')),
                ("Рекомендация", item.get('recommendation', '-')),
                (
                    "Регуляторы",
                    ", ".join(item.get('regulators', []))
                    if isinstance(item.get('regulators'), list)
                    else "-"
                ),
                (
                    "Решения",
                    ", ".join(item.get('vendors', []))
                    if isinstance(item.get('vendors'), list)
                    else "-"
                ),
                (
                    "Производители",
                    manufacturers_for_report_item(item)
                )
            ]

            for f_label, f_val in fields:

                ws.cell(row=curr_row, column=1, value=f_label).font = Font(italic=True)
                ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=4)

                ws.cell(
                    row=curr_row,
                    column=2,
                    value=f_val
                ).alignment = Alignment(wrap_text=True)

                for col in range(1, 5):
                    ws.cell(row=curr_row, column=col).border = border

                curr_row += 1

            curr_row += 1

    curr_row += 1

    # =========================
    # ROADMAP PREVIEW AT THE END
    # =========================

    ws.merge_cells(f'A{curr_row}:D{curr_row}')
    roadmap_preview_header = ws.cell(
        row=curr_row,
        column=1,
        value="ROADMAP НА 90 ДНЕЙ: ПРИОРИТЕТНЫЕ ШАГИ"
    )
    roadmap_preview_header.font = white_font
    roadmap_preview_header.fill = dark_blue_fill
    roadmap_preview_header.alignment = Alignment(horizontal='center')
    curr_row += 1

    ws.merge_cells(f'A{curr_row}:D{curr_row}')
    roadmap_hint = ws.cell(
        row=curr_row,
        column=1,
        value="Дорожная карта размещена в конце заключения: сначала оценка и риски, затем практический план действий."
    )
    roadmap_hint.alignment = Alignment(wrap_text=True)
    roadmap_hint.fill = light_blue_fill
    roadmap_hint.border = border
    curr_row += 1

    preview_headers = ["Период", "Приоритет", "Домен", "Действие"]
    for col_num, header in enumerate(preview_headers, 1):
        cell = ws.cell(row=curr_row, column=col_num, value=header)
        cell.font = white_font
        cell.fill = dark_blue_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    curr_row += 1

    for item in roadmap_items[:8]:
        ws.cell(row=curr_row, column=1, value=item["phase"]).border = border
        ws.cell(row=curr_row, column=2, value=item["priority"]).border = border
        ws.cell(row=curr_row, column=3, value=item["domain"]).border = border
        action_cell = ws.cell(row=curr_row, column=4, value=item["action"])
        action_cell.border = border
        action_cell.alignment = Alignment(wrap_text=True, vertical='top')
        if item["priority"] == "P1":
            fill = critical_fill
        elif item["priority"] == "P2":
            fill = medium_fill
        else:
            fill = light_blue_fill
        for col in range(1, 5):
            ws.cell(row=curr_row, column=col).fill = fill
        curr_row += 1

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 95
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20

    # =========================
    # ROADMAP SHEET
    # =========================

    roadmap_ws = wb.create_sheet("02 Roadmap 90 дней")
    roadmap_ws.merge_cells('A1:G1')
    roadmap_ws['A1'] = "ДОРОЖНАЯ КАРТА УЛУЧШЕНИЙ НА 90 ДНЕЙ"
    roadmap_ws['A1'].font = Font(bold=True, size=18, color="FFFFFF")
    roadmap_ws['A1'].fill = dark_blue_fill
    roadmap_ws['A1'].alignment = Alignment(horizontal='center')

    roadmap_ws.merge_cells('A3:G3')
    roadmap_ws['A3'] = f"{profile_title}. {profile_text}"
    roadmap_ws['A3'].alignment = Alignment(wrap_text=True, vertical='top')
    roadmap_ws['A3'].fill = gray_fill

    roadmap_headers = [
        "Период",
        "Приоритет",
        "Домен",
        "Что сделать",
        "Почему это важно",
        "Ответственный",
        "Сложность",
    ]

    roadmap_row = 5
    for col_num, header in enumerate(roadmap_headers, 1):
        cell = roadmap_ws.cell(row=roadmap_row, column=col_num, value=header)
        cell.font = white_font
        cell.fill = dark_blue_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    roadmap_row += 1

    priority_fills = {
        "P1": critical_fill,
        "P2": medium_fill,
        "P3": light_blue_fill,
    }

    for item in roadmap_items:
        values = [
            item["phase"],
            item["priority"],
            item["domain"],
            item["action"],
            item["rationale"],
            item["owner"],
            item["effort"],
        ]

        fill = priority_fills.get(item["priority"], gray_fill)
        for col_num, value in enumerate(values, 1):
            cell = roadmap_ws.cell(row=roadmap_row, column=col_num, value=value)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if col_num <= 2:
                cell.fill = fill

        roadmap_row += 1

    roadmap_ws.freeze_panes = "A6"
    roadmap_ws.auto_filter.ref = f"A5:G{max(roadmap_row - 1, 5)}"

    roadmap_row += 2
    roadmap_ws.merge_cells(start_row=roadmap_row, start_column=1, end_row=roadmap_row, end_column=7)
    note_cell = roadmap_ws.cell(row=roadmap_row, column=1, value=(
        "Примечание: roadmap построен с учетом размера инфраструктуры, наличия серверов, "
        "публичных сервисов, бизнес-систем и текущих средств ИБ. Для малой инфраструктуры "
        "приоритет отдается быстрым управляемым мерам, а не тяжелым enterprise-платформам."
    ))
    note_cell.alignment = Alignment(wrap_text=True, vertical='top')
    note_cell.fill = gray_fill
    note_cell.border = border

    roadmap_ws.column_dimensions['A'].width = 16
    roadmap_ws.column_dimensions['B'].width = 12
    roadmap_ws.column_dimensions['C'].width = 22
    roadmap_ws.column_dimensions['D'].width = 48
    roadmap_ws.column_dimensions['E'].width = 58
    roadmap_ws.column_dimensions['F'].width = 18
    roadmap_ws.column_dimensions['G'].width = 14

    wb.save(output)
    return output.getvalue()


def make_internal_sales_excel(c_info, results, final_score, client_report_bytes=None):
    from io import BytesIO
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    output = BytesIO()
    client_report_bytes = client_report_bytes or make_expert_excel(c_info, results, final_score)
    wb = load_workbook(BytesIO(client_report_bytes))
    ws = wb.create_sheet("03 Что продавать")

    dark_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    light_blue_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    gray_fill = PatternFill(start_color="F3F6F9", end_color="F3F6F9", fill_type="solid")
    critical_fill = PatternFill(start_color="FDE9E7", end_color="FDE9E7", fill_type="solid")
    medium_fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style='thin', color="D9D9D9"),
        right=Side(style='thin', color="D9D9D9"),
        top=Side(style='thin', color="D9D9D9"),
        bottom=Side(style='thin', color="D9D9D9")
    )

    context = build_context(results, c_info)
    domain_scores = calculate_domain_scores(results)
    rule_risks = generate_rule_based_risks(results, context)
    roadmap_items = build_contextual_roadmap(results, context, domain_scores, rule_risks)
    sales_opportunities = build_sales_opportunities(results, context, roadmap_items)
    sales_pack = build_sales_conversation_pack(
        c_info,
        results,
        context,
        roadmap_items,
        sales_opportunities
    )

    def style_sales_header(sheet, title, end_col):
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
        sheet.cell(row=1, column=1, value=title)
        sheet.cell(row=1, column=1).font = Font(bold=True, size=18, color="FFFFFF")
        sheet.cell(row=1, column=1).fill = dark_fill
        sheet.cell(row=1, column=1).alignment = Alignment(horizontal='center')

    def write_table(sheet, start_row, headers, rows, widths):
        for col_num, header in enumerate(headers, 1):
            cell = sheet.cell(row=start_row, column=col_num, value=header)
            cell.font = white_font
            cell.fill = dark_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        row_idx = start_row + 1
        for row_values in rows:
            for col_num, value in enumerate(row_values, 1):
                cell = sheet.cell(row=row_idx, column=col_num, value=value)
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                if row_idx % 2 == 0:
                    cell.fill = gray_fill
            row_idx += 1
        for col_letter, width in widths.items():
            sheet.column_dimensions[col_letter].width = width
        sheet.freeze_panes = f"A{start_row + 1}"
        sheet.auto_filter.ref = f"A{start_row}:{chr(64 + len(headers))}{max(row_idx - 1, start_row)}"
        return row_idx

    ws.merge_cells('A1:G1')
    ws['A1'] = "ВНУТРЕННИЙ SALES PLAYBOOK ПО ИТОГАМ АУДИТА"
    ws['A1'].font = Font(bold=True, size=18, color="FFFFFF")
    ws['A1'].fill = dark_fill
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A3:G3')
    ws['A3'] = (
        f"Компания: {c_info.get('Наименование компании', '-')} | "
        f"Город: {c_info.get('Город', '-')} | "
        f"Сфера: {c_info.get('Сфера деятельности', '-')} | "
        f"Зрелость ИБ: {final_score}%"
    )
    ws['A3'].alignment = Alignment(wrap_text=True)
    ws['A3'].fill = gray_fill
    ws['A3'].border = border

    info_rows = [
        ("Компания", c_info.get("Наименование компании", "-")),
        ("Город", c_info.get("Город", "-")),
        ("Сфера деятельности", c_info.get("Сфера деятельности", "-")),
        ("Сайт", c_info.get("Сайт компании", "-")),
        ("Email", c_info.get("Email", "-")),
        ("Контактное лицо", c_info.get("ФИО контактного лица", "-")),
        ("Должность", c_info.get("Должность", "-")),
        ("Телефон", c_info.get("Контактный телефон", "-")),
        ("Зрелость ИБ", f"{final_score}%"),
    ]

    ws.merge_cells('A5:G5')
    ws['A5'] = "ИНФОРМАЦИЯ О КОМПАНИИ"
    ws['A5'].font = Font(bold=True, color="FFFFFF")
    ws['A5'].fill = dark_fill
    ws['A5'].alignment = Alignment(horizontal='center')

    info_row = 6
    for label, value in info_rows:
        ws.cell(row=info_row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=info_row, column=1).fill = gray_fill
        ws.cell(row=info_row, column=1).border = border
        ws.merge_cells(start_row=info_row, start_column=2, end_row=info_row, end_column=7)
        value_cell = ws.cell(row=info_row, column=2, value=value)
        value_cell.border = border
        value_cell.alignment = Alignment(wrap_text=True, vertical='top')
        for col_num in range(3, 8):
            ws.cell(row=info_row, column=col_num).border = border
        info_row += 1

    headers = [
        "Приоритет",
        "Проблема клиента",
        "Что предложить",
        "Почему это релевантно",
        "Решения из портфеля",
        "Следующий шаг сейла",
        "Источник",
    ]
    headers.insert(5, "Дистрибьюторы")
    header_row = info_row + 1
    row = header_row
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col_num, value=header)
        cell.font = white_font
        cell.fill = dark_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    row += 1
    first_data_row = row
    priority_fills = {"P1": critical_fill, "P2": medium_fill, "P3": gray_fill}
    for item in sales_opportunities:
        values = [
            item["priority"],
            item["problem"],
            item["offer"],
            item["trigger"],
            item["vendors"],
            verified_distributors_for_vendors(item["vendors"]),
            item["next_step"],
            item.get("source", "Базовые правила"),
        ]
        fill = priority_fills.get(item["priority"], gray_fill)
        for col_num, value in enumerate(values, 1):
            cell = ws.cell(row=row, column=col_num, value=value)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if col_num == 1:
                cell.fill = fill
        row += 1

    if not sales_opportunities:
        ws.merge_cells(start_row=first_data_row, start_column=1, end_row=first_data_row, end_column=8)
        cell = ws.cell(
            row=first_data_row,
            column=1,
            value="Явных продуктовых триггеров по анкете мало. Нужно назначить короткий экспертный созвон и уточнить детали инфраструктуры."
        )
        cell.alignment = Alignment(wrap_text=True)
        cell.fill = gray_fill
        cell.border = border
        row = first_data_row + 1

    ws.freeze_panes = f"A{first_data_row}"
    ws.auto_filter.ref = f"A{header_row}:H{max(row - 1, header_row)}"
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 44
    ws.column_dimensions['D'].width = 54
    ws.column_dimensions['E'].width = 36
    ws.column_dimensions['F'].width = 34
    ws.column_dimensions['G'].width = 48
    ws.column_dimensions['H'].width = 20

    source_row = row + 2
    ws.merge_cells(start_row=source_row, start_column=1, end_row=source_row, end_column=8)
    ws.cell(row=source_row, column=1, value="ИСТОЧНИКИ РИСКОВ И РЕКОМЕНДАЦИЙ").font = white_font
    ws.cell(row=source_row, column=1).fill = dark_fill
    ws.cell(row=source_row, column=1).alignment = Alignment(horizontal='center')
    source_row += 1

    source_headers = ["Критичность", "Риск", "Рекомендация", "Источник"]
    for col_num, header in enumerate(source_headers, 1):
        cell = ws.cell(row=source_row, column=col_num, value=header)
        cell.font = white_font
        cell.fill = dark_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    source_row += 1

    risk_sources = st.session_state.get("last_report_risk_sources", [])
    if risk_sources:
        for item in risk_sources:
            values = [
                item.get("level", "-"),
                item.get("risk", "-"),
                item.get("recommendation", "-"),
                item.get("source", "-"),
            ]
            for col_num, value in enumerate(values, 1):
                cell = ws.cell(row=source_row, column=col_num, value=value)
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                if item.get("source") == "ИИ":
                    cell.fill = light_blue_fill
                elif source_row % 2 == 0:
                    cell.fill = gray_fill
            source_row += 1
    else:
        ws.merge_cells(start_row=source_row, start_column=1, end_row=source_row, end_column=4)
        cell = ws.cell(row=source_row, column=1, value="Источники рисков не зафиксированы для текущей генерации.")
        cell.fill = gray_fill
        cell.border = border
        cell.alignment = Alignment(wrap_text=True)

    ws.column_dimensions['H'].width = 1

    pains_ws = wb.create_sheet("04 Боли и гипотезы")
    style_sales_header(pains_ws, "БОЛИ КЛИЕНТА И ГИПОТЕЗЫ ПРОДАЖ", 5)
    pains_rows = [
        [
            item["priority"],
            item["pain"],
            item["evidence"],
            item["commercial_angle"],
            item["discovery_question"],
        ]
        for item in sales_pack["pains"]
    ]
    write_table(
        pains_ws,
        3,
        ["Приоритет", "Боль клиента", "Факт из анкеты", "Коммерческий заход", "Вопрос для подтверждения"],
        pains_rows,
        {"A": 12, "B": 42, "C": 42, "D": 44, "E": 52}
    )

    script_ws = wb.create_sheet("05 Сценарий звонка")
    style_sales_header(script_ws, "СЦЕНАРИЙ ПЕРВОГО ЗВОНКА", 3)
    script_rows = [
        [idx, item["stage"], item["talk_track"]]
        for idx, item in enumerate(sales_pack["call_script"], start=1)
    ]
    write_table(
        script_ws,
        3,
        ["#", "Этап", "Что сказать"],
        script_rows,
        {"A": 8, "B": 28, "C": 110}
    )

    questions_ws = wb.create_sheet("06 Вопросы")
    style_sales_header(questions_ws, "ВОПРОСЫ ДЛЯ УТОЧНЕНИЯ НА ВСТРЕЧЕ", 2)
    write_table(
        questions_ws,
        3,
        ["Тема", "Вопрос"],
        [[topic, question] for topic, question in sales_pack["questions"]],
        {"A": 28, "B": 110}
    )

    objections_ws = wb.create_sheet("07 Возражения")
    style_sales_header(objections_ws, "ТИПОВЫЕ ВОЗРАЖЕНИЯ И ОТРАБОТКА", 2)
    write_table(
        objections_ws,
        3,
        ["Возражение", "Как отвечать"],
        [[objection, answer] for objection, answer in sales_pack["objections"]],
        {"A": 45, "B": 105}
    )

    next_ws = wb.create_sheet("08 Следующие шаги")
    style_sales_header(next_ws, "РЕКОМЕНДУЕМЫЕ СЛЕДУЮЩИЕ ШАГИ", 4)
    next_rows = [
        [
            item["priority"],
            item["offer"],
            item["step"],
            item["success_criteria"],
        ]
        for item in sales_pack["next_steps"]
    ]
    write_table(
        next_ws,
        3,
        ["Приоритет", "Предложение", "Действие сейла", "Критерий успеха"],
        next_rows,
        {"A": 12, "B": 42, "C": 60, "D": 70}
    )

    answers_ws = wb.create_sheet("09 Заполненная анкета")
    answers_ws.merge_cells('A1:B1')
    answers_ws['A1'] = "ЗАПОЛНЕННАЯ АНКЕТА / ДАННЫЕ ДЛЯ СЕЙЛА И ПРЕСЕЙЛА"
    answers_ws['A1'].font = Font(bold=True, size=18, color="FFFFFF")
    answers_ws['A1'].fill = dark_fill
    answers_ws['A1'].alignment = Alignment(horizontal='center')

    answers_ws.cell(row=3, column=1, value="Поле").font = white_font
    answers_ws.cell(row=3, column=2, value="Значение").font = white_font
    for col_num in range(1, 3):
        cell = answers_ws.cell(row=3, column=col_num)
        cell.fill = dark_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center')

    answer_row = 4
    for key, value in results.items():
        if str(key).startswith("_"):
            continue
        answers_ws.cell(row=answer_row, column=1, value=key).border = border
        value_cell = answers_ws.cell(row=answer_row, column=2, value=str(value))
        value_cell.border = border
        value_cell.alignment = Alignment(wrap_text=True, vertical='top')
        if answer_row % 2 == 0:
            answers_ws.cell(row=answer_row, column=1).fill = gray_fill
            answers_ws.cell(row=answer_row, column=2).fill = gray_fill
        answer_row += 1

    answers_ws.freeze_panes = "A4"
    answers_ws.auto_filter.ref = f"A3:B{max(answer_row - 1, 3)}"
    answers_ws.column_dimensions['A'].width = 42
    answers_ws.column_dimensions['B'].width = 100

    wb.save(output)
    return output.getvalue(), sales_opportunities


def build_report_results(
    data,
    main_speed,
    back_speed,
    total_arm,
    ap_cnt,
    selected_routing,
    ngfw_vendor,
    phys_count,
    virt_count,
    v_n_b
):
    results = data.copy()
    results.update({
        "Интернет канал (осн)": f"{main_speed} Mbit/s",
        "Резервный канал": f"{back_speed} Mbit/s",
        "_main_speed": main_speed,
        "_back_speed": back_speed,
        "_user_count": total_arm,
        "WiFi Точки": ap_cnt,
        "WiFi Контроллер": data.get('Wi-Fi Контроллер', "Нет"),
        "Маршрутизация": ", ".join(selected_routing) if selected_routing else "Нет",
        "NGFW": ngfw_vendor if ngfw_vendor else "Нет",
        "Серверы (физ)": phys_count,
        "Серверы (вирт)": virt_count,
        "Резервное копирование": v_n_b if v_n_b else "Нет",
    })

    results["MFA"] = results.get("Блок 2. MFA", "Нет")
    results["SIEM"] = results.get("Блок 2. SIEM", "Нет")
    results["WAF"] = results.get("Блок 2. WAF", "Нет")
    results["Anti-DDoS"] = results.get("Блок 2. Anti-DDoS", "Нет")
    results["EPP"] = results.get("Блок 2. EPP", "Нет")
    results["Антивирус"] = results.get("Блок 2. EPP", "Нет")
    results["EDR"] = results.get("Блок 2. EDR", "Нет")
    results["XDR"] = results.get("Блок 2. XDR", "Нет")
    results["MDR"] = results.get("Блок 2. MDR", "Нет")
    results["DLP"] = results.get("Блок 2. DLP", "Нет")
    results["Mail Security"] = results.get("Блок 2. Mail Security", "Нет")
    results["CASB"] = results.get("Блок 2. CASB", "Нет")
    results["IDS/IPS"] = results.get("Блок 2. IDS/IPS", "Нет")
    results["NAC"] = results.get("Блок 2. NAC", "Нет")
    results["ZTNA"] = results.get("Блок 2. ZTNA", "Нет")
    results["SAST"] = results.get("Блок 2. SAST", "Нет")
    results["DAST"] = results.get("Блок 2. DAST", "Нет")
    results["IAM"] = results.get("Блок 2. IAM", "Нет")
    results["PAM"] = results.get("Блок 2. PAM", "Нет")
    results["SOAR"] = results.get("Блок 2. SOAR", "Нет")
    results["NAD"] = results.get("Блок 2. NAD", "Нет")
    results["Patch Management"] = results.get("Блок 2. Patch Management", "Нет")
    results["Виртуализация"] = "Да" if virt_count > 0 else "Нет"
    results["СХД"] = "Да" if any(str(key).startswith("1.4.") for key in results) else "Нет"
    results["Резервный канал"] = f"{back_speed} Mbit/s" if back_speed > 0 else "Нет"
    return results


def _status_label(score_value):
    if score_value >= 80:
        return "Сильный"
    if score_value >= 50:
        return "Средний"
    return "Слабый"


def _quick_win_signals(results):
    signals = []

    if results.get("MFA") == "Нет":
        signals.append(("critical", "MFA", "Внедрение MFA обычно быстрее всего снижает риск компрометации учетных записей."))

    if results.get("SIEM") == "Нет":
        signals.append(("warn", "SIEM / SOC", "Нет централизованного мониторинга: инциденты сложнее обнаруживать и расследовать."))

    if results.get("EDR") == "Нет":
        signals.append(("warn", "Endpoint", "Без EDR/XDR защита рабочих мест остается слабее против сложных атак."))

    if results.get("Резервное копирование") == "Нет":
        signals.append(("critical", "Backup", "Не указан backup-контур: это критично для устойчивости к ransomware и сбоям."))

    if results.get("Patch Management") == "Нет":
        signals.append(("warn", "Patch Management", "Без централизованных обновлений растет риск эксплуатации известных уязвимостей."))

    legacy_arm = results.get("ОС АРМ (Windows XP/Vista/7/8)", 0)
    legacy_srv = results.get("ОС Сервера (Windows Server 2008/2012 R2)", 0)
    if legacy_arm or legacy_srv:
        signals.append(("critical", "Legacy OS", "Обнаружены устаревшие ОС: нужен план миграции или изоляции."))

    return signals[:5]


def section_status(enabled, complete):
    if not enabled:
        return "disabled"
    return "complete" if complete else "missing"


def unpack_section_status(section_item):
    if len(section_item) == 3:
        return section_item

    step, status = section_item
    return step, status, ""


def calculate_audit_readiness(client_info, validation_errors, section_statuses):
    required_fields = [
        client_info.get('Город'),
        client_info.get('Наименование компании'),
        client_info.get('Сфера деятельности'),
        client_info.get('Сайт компании'),
        client_info.get('Email'),
        client_info.get('ФИО контактного лица'),
        client_info.get('Должность'),
        client_info.get('Контактный телефон'),
    ]

    filled_required = sum(1 for value in required_fields if value)
    required_readiness = (filled_required / len(required_fields)) * 40

    active_sections = []
    for section_item in section_statuses:
        _, status, _ = unpack_section_status(section_item)
        if status != "disabled":
            active_sections.append(status)
    completed_sections = sum(
        1 for status in active_sections
        if status == "complete"
    )
    section_readiness = (
        (completed_sections / len(active_sections)) * 45
        if active_sections
        else 0
    )
    validation_readiness = 15 if not validation_errors else 0

    return int(round(
        min(100, required_readiness + section_readiness + validation_readiness)
    ))


def render_audit_cockpit(client_info, results, validation_errors, final_score, it_maturity_score, section_statuses):
    readiness = calculate_audit_readiness(
        client_info,
        validation_errors,
        section_statuses
    )
    domain_scores = calculate_domain_scores(results)
    risk_signals = _quick_win_signals(results)
    maturity, maturity_icon = get_maturity_level(final_score)
    it_maturity, it_maturity_icon = get_maturity_level(it_maturity_score)

    st.sidebar.markdown("### Навигатор аудита")
    st.sidebar.progress(readiness, text=f"Готовность анкеты: {readiness}%")
    st.sidebar.metric("Зрелость ИБ", f"{final_score}%")
    st.sidebar.caption(f"{maturity_icon} {maturity}")
    st.sidebar.metric("Зрелость ИТ", f"{it_maturity_score}%")
    st.sidebar.caption(f"{it_maturity_icon} {it_maturity}")
    st.sidebar.metric("Ошибки заполнения", len(validation_errors))

    st.sidebar.markdown("#### Разделы")
    status_styles = {
        "complete": ("green", "заполнен"),
        "missing": ("red", "не заполнен"),
        "disabled": ("gray", "не включен"),
    }

    for section_item in section_statuses:
        step, status, anchor_id = unpack_section_status(section_item)
        dot_class, title = status_styles.get(status, status_styles["disabled"])
        anchor_attr = f' href="#{html.escape(anchor_id)}"' if anchor_id else ""
        st.sidebar.markdown(
            f'<a class="sidebar-step sidebar-step-link"{anchor_attr} title="{html.escape(title)}">'
            f'<span class="sidebar-dot {dot_class}"></span><span>{html.escape(step)}</span>'
            f'</a>',
            unsafe_allow_html=True
        )

    if validation_errors:
        st.sidebar.markdown("#### Требует внимания")
        for err in list(dict.fromkeys(validation_errors))[:6]:
            anchor_id = get_error_anchor(err)
            st.sidebar.markdown(
                f'<a class="sidebar-error-link" href="#{html.escape(anchor_id)}">'
                f'<span>Исправить</span>{html.escape(err)}'
                f'</a>',
                unsafe_allow_html=True
            )


def render_analysis_preview(results, final_score):
    domain_scores = calculate_domain_scores(results)
    risk_signals = _quick_win_signals(results)
    maturity, maturity_icon = get_maturity_level(final_score)

    st.markdown("### Сводка аудита")
    metric_cols = st.columns(4)
    metric_values = [
        ("Зрелость ИБ", f"{final_score}%", f"{maturity_icon} {maturity}"),
        ("Доменов", str(len(domain_scores)), "в текущей оценке"),
        ("Рисков", str(len(risk_signals)), "быстрые сигналы"),
        ("Статус", "Предпросмотр", "перед XLSX"),
    ]

    for col, (label, value, hint) in zip(metric_cols, metric_values):
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="label">{html.escape(label)}</div>
                <div class="value">{html.escape(value)}</div>
                <div class="hint">{html.escape(hint)}</div>
            </div>
            """, unsafe_allow_html=True)

    domain_cards = ''.join(
        f'<div class="domain-card">'
        f'<strong>{html.escape(domain)}</strong>'
        f'<div class="domain-score">{score_value}%</div>'
        f'<div class="hint">{html.escape(_status_label(score_value))}</div>'
        f'</div>'
        for domain, score_value in domain_scores.items()
    )
    st.markdown(f'<div class="domain-row">{domain_cards}</div>', unsafe_allow_html=True)

    if risk_signals:
        st.markdown("#### Быстрые улучшения")
        for kind, title, body in risk_signals:
            st.markdown(f"""
            <div class="risk-chip {html.escape(kind)}">
                <strong>{html.escape(title)}</strong><br>
                {html.escape(body)}
            </div>
            """, unsafe_allow_html=True)


def render_analysis_teaser(results, final_score, validation_errors, section_statuses, client_info):
    domain_scores = calculate_domain_scores(results)
    risk_signals = _quick_win_signals(results)
    maturity, maturity_icon = get_maturity_level(final_score)
    readiness = calculate_audit_readiness(
        client_info,
        validation_errors,
        section_statuses
    )

    weak_domains = sorted(
        domain_scores.items(),
        key=lambda item: item[1]
    )[:3]
    top_signals = risk_signals[:2]

    if top_signals:
        signal_html = "".join(
            f'<div class="mini-signal {html.escape(kind)}">'
            f'<strong>{html.escape(title)}</strong><br>'
            f'{html.escape(body)}'
            f'</div>'
            for kind, title, body in top_signals
        )
    else:
        signal_html = (
            '<div class="mini-signal">'
            '<strong>Быстрые улучшения появятся здесь</strong><br>'
            'Заполните ключевые блоки, и система покажет первые практические рекомендации.'
            '</div>'
        )

    domain_html = "".join(
        f'<div class="mini-domain">'
        f'<span>{html.escape(domain)}</span>'
        f'<strong>{score_value}%</strong>'
        f'</div>'
        for domain, score_value in weak_domains
    )

    st.markdown(f"""
    <div class="analysis-teaser">
        <div class="analysis-teaser-head">
            <div>
                <div class="analysis-teaser-title">Предварительная аналитика уже собирается</div>
                <div class="analysis-teaser-copy">
                    По мере заполнения анкеты здесь появляется ранняя оценка защиты,
                    слабые домены и быстрые улучшения. Полную сводку можно раскрыть ниже.
                </div>
            </div>
            <div class="analysis-pill">обновляется автоматически</div>
        </div>
        <div class="teaser-grid">
            <div class="teaser-card">
                <div class="label">Зрелость ИБ</div>
                <div class="value">{final_score}%</div>
                <div class="hint">{html.escape(maturity_icon)} {html.escape(maturity)}</div>
            </div>
            <div class="teaser-card">
                <div class="label">Готовность анкеты</div>
                <div class="value">{readiness}%</div>
                <div class="hint">до формирования XLSX</div>
            </div>
            <div class="teaser-card">
                <div class="label">Быстрые сигналы</div>
                <div class="value">{len(risk_signals)}</div>
                <div class="hint">первые практические подсказки</div>
            </div>
        </div>
        <div class="teaser-columns">
            <div>
                <div class="label">Что система уже видит</div>
                {signal_html}
            </div>
            <div>
                <div class="label">Домены, требующие внимания</div>
                {domain_html}
            </div>
        </div>
        <div class="analysis-teaser-note">
            Раскройте полный блок ниже, чтобы увидеть все домены, статусы и быстрые улучшения.
        </div>
    </div>
    """, unsafe_allow_html=True)

# --- ИНИЦИАЛИЗАЦИЯ И СТЕК СОСТОЯНИЙ (в самом начале финального блока) ---
if "generation_state" not in st.session_state:
    st.session_state.generation_state = "idle"  # Может быть: idle, preparing, heavy_ai, finalized
if "cached_report_bytes" not in st.session_state:
    st.session_state.cached_report_bytes = None
if "cached_sales_report_bytes" not in st.session_state:
    st.session_state.cached_sales_report_bytes = None
if "telegram_status" not in st.session_state:
    st.session_state.telegram_status = ""
if "generation_attempt_started_at" not in st.session_state:
    st.session_state.generation_attempt_started_at = None
if "report_shortened_last" not in st.session_state:
    st.session_state.report_shortened_last = False
if "last_report_risk_sources" not in st.session_state:
    st.session_state.last_report_risk_sources = []
if "telegram_generation_started_sent" not in st.session_state:
    st.session_state.telegram_generation_started_sent = False

render_generation_guard(
    st.session_state.generation_state in {"preparing", "heavy_ai"}
)

preview_results = build_report_results(
    data,
    main_speed,
    back_speed,
    total_arm,
    ap_cnt,
    selected_routing,
    ngfw_vendor,
    phys_count,
    virt_count,
    v_n_b
)
audit_started = any(str(value).strip() for value in client_info.values()) or total_arm > 0 or any([
    net_active,
    server_active,
    storage_active,
    is_active,
    enable_security,
    dev_active,
])
security_controls = [
    (epp, epp_v, 8),
    (edr, edr_v, 12),
    (xdr, xdr_v, 8),
    (mdr, mdr_v, 8),
    (dlp, dlp_v, 6),
    (mail_sec, mail_v, 5),
    (casb, casb_v, 4),
    (waf, waf_v, 6),
    (ddos, ddos_v, 4),
    (ids, ids_v, 6),
    (nac, nac_v, 5),
    (ztna, ztna_v, 5),
    (sast, sast_v, 5),
    (dast, dast_v, 5),
    (iam, iam_v, 5),
    (mfa, mfa_v, 10),
    (pam, pam_v, 8),
    (siem, siem_v, 10),
    (soar, soar_v, 4),
    (vuln, vuln_v, 5),
    (patch, patch_v, 6),
    (nad, nad_v, 5),
]
preview_score = calculate_weighted_security_score(enable_security, security_controls)

def has_validation_error(*markers):
    return any(
        any(marker in err for marker in markers)
        for err in validation_errors
    )


general_complete = all([
    client_info.get('Город'),
    client_info.get('Наименование компании'),
    client_info.get('Сфера деятельности'),
    client_info.get('Сайт компании'),
    client_info.get('Email'),
    client_info.get('ФИО контактного лица'),
    client_info.get('Должность'),
    client_info.get('Контактный телефон'),
])
endpoint_complete = total_arm > 0 and bool(selected_os_arm) and sum_os_arm == total_arm
network_complete = net_active and bool(selected_routing) and not has_validation_error(
    "маршрутизации",
    "маршрутизаторов",
    "коммутаторов",
    "Core-уровня",
    "уровня распределения",
    "уровня доступа",
    "Wi-Fi",
    "NGFW",
)
server_complete = server_active and (phys_count > 0 or virt_count > 0) and not has_validation_error(
    "серверов",
    "хостов",
    "резервного копирования",
)
storage_complete = storage_active and not has_validation_error("СХД", "RAID")
systems_complete = is_active and not has_validation_error("название и версию", "версию")
security_complete = enable_security and not has_validation_error("производитель")
web_complete = web_active
dev_complete = dev_active and not has_validation_error("разработчиков", "языки разработки")

it_maturity_score = calculate_it_maturity_score(
    total_arm,
    selected_os_arm,
    sum_os_arm,
    net_active,
    main_speed,
    back_speed,
    selected_routing,
    ap_cnt,
    ngfw_vendor,
    server_active,
    phys_count,
    virt_count,
    selected_virt_sys,
    v_n_b,
    storage_active,
    st_media_sel,
    cnt_hdd,
    cnt_ssd,
    raid_selected,
    is_active,
    web_active,
    dev_active,
    dev_count,
    sel_langs,
    cicd_active,
)

section_statuses = [
    ("Компания", section_status(True, general_complete), "section-company"),
    ("Конечные точки", section_status(True, endpoint_complete), "section-endpoints"),
    ("Сеть", section_status(net_active, network_complete), "section-network"),
    ("Серверы", section_status(server_active, server_complete), "section-servers"),
    ("СХД", section_status(storage_active, storage_complete), "section-storage"),
    ("ИС", section_status(is_active, systems_complete), "section-systems"),
    ("ИБ", section_status(enable_security, security_complete), "section-security"),
    ("Веб", section_status(web_active, web_complete), "section-web"),
    ("Разработка", section_status(dev_active, dev_complete), "section-dev"),
]

render_audit_cockpit(
    client_info,
    preview_results,
    validation_errors,
    preview_score,
    it_maturity_score,
    section_statuses
)

render_analysis_teaser(
    preview_results,
    preview_score,
    validation_errors,
    section_statuses,
    client_info
)

with st.expander("Открыть полный предварительный анализ", expanded=False):
    render_analysis_preview(preview_results, preview_score)

# --- ФИНАЛ ---
st.divider()
if validation_errors:
    render_validation_summary(validation_errors)

# КНОПКА ЗАПУСКА ПРОЦЕССА
# Она активна только тогда, когда процесс еще не запущен
st.markdown("""
Нажимая «Сформировать экспертный отчет», вы даете согласие
на обработку персональных данных в соответствии с
<a href="https://drive.google.com/file/d/1ypEIH9_ePGo3elkR2ifLFBulD5CAFOfs/view?usp=sharing" target="_blank">
Политикой конфиденциальности
</a>.
""", unsafe_allow_html=True)
if st.session_state.generation_state == "idle":
    if st.button("Сформировать экспертный отчет", disabled=len(validation_errors) > 0, type="primary"):
        st.session_state.telegram_generation_started_sent = False
        render_generation_guard(True)
        alert_placeholder = st.empty()
        console_placeholder = st.empty()
        progress_bar = st.progress(0)

        alert_placeholder.markdown(
            """
            <div class="analysis-status-panel">
                <div class="analysis-status-title">Формируется экспертный отчет</div>
                Выполняется нормализация данных, расчет зрелости и сборка XLSX. Это может занять до 4 минут.
                <div class="page-lock-note">Не закрывайте и не обновляйте страницу до завершения формирования.</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        steps = [

            "Инициализация ядра аудита...",
            "Проверка обязательных полей...",
            "Нормализация инфраструктурных данных...",
            "Анализ сетевого периметра...",
            "Анализ защищенности конечных точек...",
            "Проверка устойчивости резервного копирования...",
            "Расчет зрелости киберзащиты...",
            "Построение доменов безопасности...",
            "Глубокий анализ рисков...",
            "Формирование управленческого резюме...",
            "Генерация XLSX отчета...",
            "Финализация артефактов..."
        ]

        active_logs = []

        progress = 0

        for step in steps:

            active_logs.append(
                f"[{time.strftime('%H:%M:%S')}] {step}"
            )

            if len(active_logs) > 4:
                active_logs.pop(0)

            console_placeholder.markdown(
                '<div class="analysis-log">' +
                "".join([f'<div>▶ {line}</div>' for line in active_logs]) +
                '</div>',
                unsafe_allow_html=True
            )

            progress += random.randint(5, 9)

            progress_bar.progress(min(progress, 88))

            time.sleep(random.uniform(0.7, 1.4))


        
        # При клике мы просто меняем статус в памяти на "подготовка" и мгновенно перезапускаем страницу
        st.session_state.generation_state = "preparing"
        st.rerun()

# --- СЦЕНАРИЙ 1: ЭКРАН ОЖИДАНИЯ С ФАКТАМИ ИБ (Показывается СРАЗУ же после клика) ---
if st.session_state.generation_state == "preparing":

    # 1. Сразу жестко выводим на экран поле логов и факты информационной безопасности
    st.markdown("#### 🛠️ Ход выполнения анализа:")
    render_generation_live_panel("Подготовка экспертного анализа", active_step=1)

    # Имитируем лог-систему, как вы просили
    st.info("⚙️ `[СИСТЕМА]`: Инициализация аналитического ядра Khalil Consulting v11.01...")
    st.success("⚙️ `[МАТРИЦА]`: Агрегация параметров ИТ-инфраструктуры успешно завершена.")
    
    st.markdown("---")
    st.markdown("#### Полезные факты и рекомендации по ИБ:")
    
    # Выводим на экран массив фактов в красивом поле, который пользователь будет читать все 3 минуты
    st.markdown("""
    <div class="facts-panel">
        <p><strong>Многофакторная аутентификация:</strong> внедрение MFA блокирует до 99.9% автоматизированных атак на корпоративные учетные записи.</p>
        <p><strong>Защита рабочих мест:</strong> обычного антивируса (EPP) в 2026 году уже недостаточно. Решения класса EDR/XDR критически необходимы для выявления скрытых бесфайловых угроз.</p>
        <p><strong>Безопасность архивов:</strong> резервные копии должны быть изолированы от основной сети. Принцип Immutable Backup снижает риск шифрования бэкапов.</p>
        <p><strong>Сетевой периметр:</strong> сетевая сегментация, VLAN и Zero Trust помогают остановить распространение атаки внутри компании.</p>
        <p><strong>Человеческий фактор:</strong> значительная часть успешных кибератак начинается с фишингового письма. Регулярное обучение снижает этот риск.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Делаем маленькую паузу в 1.5 секунды, чтобы Streamlit успел железно отправить этот интерфейс в браузер клиента
    time.sleep(1.5)
    
    # Меняем статус на "Запуск тяжелого ИИ" и перезапускаем страницу. 
    # Теперь этот красивый экран останется висеть в браузере, пока ИИ думает!
    st.session_state.generation_state = "heavy_ai"
    st.rerun()

# --- СЦЕНАРИЙ 2: ЗАПУСК ТЯЖЕЛОГО ИИ И СБОРКИ EXCEL ---
if st.session_state.generation_state == "heavy_ai":
    if st.session_state.generation_attempt_started_at is None:
        st.session_state.generation_attempt_started_at = time.time()
    elif time.time() - st.session_state.generation_attempt_started_at > 300:
        st.session_state.generation_state = "idle"
        st.session_state.generation_attempt_started_at = None
        st.error("Формирование отчета было сброшено по таймауту. Запустите формирование еще раз.")
        st.stop()

    render_generation_live_panel("Идет глубокий анализ и сборка отчета", active_step=4)

    if not st.session_state.telegram_generation_started_sent:
        st.session_state.telegram_status = send_internal_telegram_message(
            build_telegram_generation_started_text(client_info, preview_score)
        )
        if st.session_state.telegram_status == "ok":
            st.session_state.telegram_generation_started_sent = True

    # Этот текст и анимация будут гореть параллельно с фактами сверху
    with st.spinner("Производится глубокий анализ рисков..."):
        try:
            # Подготовка данных перед передачей
            results = build_report_results(
                data,
                main_speed,
                back_speed,
                total_arm,
                ap_cnt,
                selected_routing,
                ngfw_vendor,
                phys_count,
                virt_count,
                v_n_b
            )
            f_score = preview_score

            st.session_state.ai_last_error = ""
            st.session_state.ai_model_used = ""
            st.session_state.ai_used_in_last_report = False
            st.session_state.ai_audit_narrative = {}

            # Запуск экспертного анализа и сборки клиентского XLSX.
            report_bytes = make_expert_excel(client_info, results, f_score)
            ai_report_ready = (
                bool(st.session_state.get("ai_used_in_last_report"))
                and not st.session_state.get("ai_last_error")
            )
            st.session_state.report_shortened_last = False
            if not ai_report_ready:
                st.session_state.telegram_status = send_internal_telegram_message(
                    build_telegram_ai_failure_text(
                        client_info,
                        f_score,
                        st.session_state.get("ai_last_error", "AI analysis did not return recommendations")
                    )
                )

            sales_report_bytes, telegram_sales = make_internal_sales_excel(
                client_info,
                results,
                f_score,
                report_bytes
            )
            st.session_state.cached_report_bytes = report_bytes
            st.session_state.cached_sales_report_bytes = sales_report_bytes
        except Exception as exc:
            st.session_state.telegram_status = send_internal_telegram_message(
                build_telegram_generation_error_text(
                    client_info,
                    preview_score,
                    redact_secret(exc, TOKEN)
                )
            )
            st.session_state.generation_state = "idle"
            st.session_state.generation_attempt_started_at = None
            st.error("Не удалось сформировать отчет. Попробуйте повторить позже.")
            st.stop()

    # Тихо отправляем в ТГ без создания задержек на экране
    st.session_state.telegram_status = ""
    if TOKEN and CHAT_ID:
        try:
            sales_lines = []
            for idx, item in enumerate(telegram_sales[:3], start=1):
                sales_lines.append(
                    f"{idx}. {item['offer']} | {item['vendors']}"
                )
            sales_digest = "\n".join(sales_lines) if sales_lines else "Нет явных продуктовых триггеров, нужен экспертный разбор."

            telegram_text = build_telegram_lead_text(
                client_info,
                f_score,
                sales_digest
            )
            telegram_send_node(
                TOKEN,
                "sendMessage",
                {"chat_id": CHAT_ID, "text": telegram_text},
                timeout_seconds=8
            )

            telegram_send_node(
                TOKEN,
                "sendDocument",
                {
                    "chat_id": CHAT_ID,
                    "caption": f"Sales playbook с клиентским отчетом: {client_info['Наименование компании']}"
                },
                files=[{
                    "field": "document",
                    "filename": f"Sales_Playbook_{client_info['Наименование компании']}.xlsx",
                    "bytes": sales_report_bytes,
                    "suffix": ".xlsx",
                }],
                timeout_seconds=15
            )
            st.session_state.telegram_status = "ok"
        except Exception as exc:
            st.session_state.telegram_status = f"Telegram не отправлен: {redact_secret(exc, TOKEN)}"
    else:
        st.session_state.telegram_status = "Telegram не отправлен: не найдены TELEGRAM_TOKEN или TELEGRAM_CHAT_ID."

    # Переключаем статус в финал
    st.session_state.generation_state = "finalized"
    st.session_state.generation_attempt_started_at = None
    st.rerun()

# --- СЦЕНАРИЙ 3: НЕУДАЧНАЯ СБОРКА БЕЗ КЛИЕНТСКОГО ОТЧЕТА ---
if st.session_state.generation_state == "ai_failed":
    st.session_state.generation_state = "idle"
    st.session_state.generation_attempt_started_at = None
    st.rerun()

# --- СЦЕНАРИЙ 4: ВЫВОД ГОТОВОГО РЕЗУЛЬТАТА ---
if st.session_state.generation_state == "finalized":

    st.success("🎉 Экспертный отчет успешно сформирован и проверен системой контроля качества Khalil Consulting!")

    st.download_button(
        label="Скачать готовый экспертный отчет (XLSX)",
        data=st.session_state.cached_report_bytes,
        file_name=f"Audit_Khalil_{client_info['Наименование компании']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    # Кнопка для сброса состояния, если пользователь захочет перегенерировать отчет
    if st.button("Сформировать новый отчет"):
        st.session_state.generation_state = "idle"
        st.session_state.cached_report_bytes = None
        st.session_state.cached_sales_report_bytes = None
        st.session_state.telegram_status = ""
        st.session_state.ai_last_error = ""
        st.session_state.report_shortened_last = False
        st.session_state.generation_attempt_started_at = None
        st.session_state.telegram_generation_started_sent = False
        st.rerun()

st.info("Khalil Audit System v11.01 | by Ivan Rudoy | Алматы 2026")
