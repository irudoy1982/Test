from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import crm_store
import crm_admin
import crm_assets


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
    generated = crm_admin._hash_admin_password("a-strong-test-password")
    assert_true(
        crm_admin._verify_admin_password("a-strong-test-password", generated),
        "Generated admin password hash cannot be verified",
    )


def test_managed_asset_validation() -> None:
    logo = crm_assets.validate_logo((ROOT / "logo.png").read_bytes(), "logo.png")
    presentation = crm_assets.validate_presentation_template(
        (ROOT / "static" / "audit_presentation_khalil.pptx").read_bytes(),
        "audit_presentation_khalil.pptx",
    )
    portfolio = crm_assets.validate_vendor_matrix(
        (ROOT / "vendor_matrix_detailed.xlsx").read_bytes(),
        "vendor_matrix_detailed.xlsx",
    )
    assert_true(logo.ok, logo.message)
    assert_true(presentation.ok, presentation.message)
    assert_true(portfolio.ok, portfolio.message)


def test_private_asset_download_route() -> None:
    calls = []

    class FakeResponse:
        status_code = 200
        text = ""

        def __init__(self, content, payload=None):
            self.content = content
            self.payload = payload

        def json(self):
            return self.payload

    def fake_request(method, url, **kwargs):
        calls.append((method, url, kwargs))
        if "/rest/v1/admin_assets" in url:
            return FakeResponse(
                b"metadata",
                [{"asset_key": "logo", "object_path": "published/logo/version_logo.png"}],
            )
        return FakeResponse(b"asset")

    original_request = crm_store.requests.request
    crm_store.requests.request = fake_request
    try:
        store = crm_store.SupabaseCrmStore("https://example.supabase.co", "service-key")
        assert_true(store.download_asset("logo") == b"asset", "Private asset was not downloaded")
    finally:
        crm_store.requests.request = original_request
    assert_true(
        any(
            "/storage/v1/object/authenticated/audit-admin-assets/published/logo/version_logo.png"
            in call[1]
            for call in calls
        ),
        "Private Supabase asset route is incorrect",
    )


def main() -> None:
    tests = [
        test_runtime_defaults_and_validation,
        test_normalized_lead_payload,
        test_amo_domain_validation,
        test_admin_password_verification,
        test_managed_asset_validation,
        test_private_asset_download_route,
    ]
    for test in tests:
        test()
        print(f"OK {test.__name__}")
    print("CRM ADMIN TEST PASSED")


if __name__ == "__main__":
    main()
