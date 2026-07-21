from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import crm_store
import crm_admin


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def test_runtime_defaults_and_validation() -> None:
    defaults = crm_store.normalize_runtime_settings({})
    assert_true(defaults["active_provider"] == "off", "CRM must default to off")
    assert_true(defaults["customer_delivery_format"] == "pptx", "PPTX must remain the default")
    assert_true(defaults["telegram_diagnostics_enabled"] is True, "Diagnostics default changed")
    invalid = crm_store.normalize_runtime_settings(
        {"active_provider": "unknown", "customer_delivery_format": "pdf"}
    )
    assert_true(invalid["active_provider"] == "off", "Unknown provider must be rejected")
    assert_true(invalid["customer_delivery_format"] == "pptx", "Unknown format must be rejected")
    both = crm_store.normalize_runtime_settings({"customer_delivery_format": "both"})
    assert_true(both["customer_delivery_format"] == "both", "Combined customer format was rejected")


def test_normalized_lead_payload() -> None:
    payload = crm_store.build_normalized_lead_payload(
        {
            "Наименование компании": "Demo LLP",
            "Сфера деятельности": "Ритейл",
            "Email": " SALES@EXAMPLE.KZ ",
            "Контактный телефон": "+7 (777) 123-45-67",
        },
        150,
        -1,
        "Test",
    )
    assert_true(payload["schema"] == "audit-crm-lead-v1", "Unexpected CRM schema")
    assert_true(payload["email"] == "sales@example.kz", "Email was not normalized")
    assert_true(payload["phone"] == "+77771234567", "Phone was not normalized")
    assert_true(payload["security_maturity"] == 100, "Security maturity was not capped")
    assert_true(payload["it_maturity"] == 0, "IT maturity was not bounded")


def test_amo_domain_validation() -> None:
    assert_true(
        crm_store.normalize_amo_domain("https://demo.amocrm.ru") == "demo.amocrm.ru",
        "Valid amoCRM domain was rejected",
    )
    try:
        crm_store.normalize_amo_domain("https://example.com/steal")
    except crm_store.CrmConfigurationError:
        pass
    else:
        raise AssertionError("Foreign amoCRM domain must be rejected")


def test_admin_password_verification() -> None:
    configured = (
        "pbkdf2_sha256$1$MDEyMzQ1Njc4OWFiY2RlZg==$"
        "bqbm_HAVeKUOuXe8pIQvt3Pbpek_GA8FXLXfptrrbbw="
    )
    assert_true(
        crm_admin._verify_admin_password("test-password", configured),
        "PBKDF2 admin password was rejected",
    )
    assert_true(
        not crm_admin._verify_admin_password("wrong-password", configured),
        "Wrong admin password was accepted",
    )


def main() -> None:
    tests = [
        test_runtime_defaults_and_validation,
        test_normalized_lead_payload,
        test_amo_domain_validation,
        test_admin_password_verification,
    ]
    for test in tests:
        test()
        print(f"OK {test.__name__}")
    print("CRM ADMIN TEST PASSED")


if __name__ == "__main__":
    main()
