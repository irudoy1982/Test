# Versioning

## Release lanes

- `v12.33`: current stable production release in KhalilAudit and BTGAudit.
- `vX3-dev`: current development lane in Test.
- `vX3-rcN`: release candidate after Test validation; bug fixes only.
- `vX3`: stable release after explicit production approval.

## Rules

- Work on new features only in `Test`.
- Do not modify `KhalilAudit` or `BTGAudit` during experiments.
- Promote to production only after the user explicitly approves it.
- Customer-facing changes go to `CHANGELOG_CUSTOMER.md`.
- Internal changes can be tracked in commits and technical notes, but must not be shown to customers.
- Minor releases increment the number after the point: `v12.31`, `v12.32`, `v12.33`.
- Development and release-candidate suffixes are retained: `v12.1-dev`, `v12.1-rc1`.
- Major version `13` is written as `X3`: `vX3-dev`, `vX3-rc1`, `vX3`.
- The digit `3` in minor versions is not replaced: `v12.3` is valid.

## Completed v12 scope

1. Branded customer presentation and separate sales playbook.
2. AI-first recommendations with deterministic fallback and fact checks.
3. Draft download, restoration, and handoff by shareable link.
4. Sector-aware Kazakhstan regulatory context.
5. Automated smoke and deep tests for reports, presentations, vendors, and core questionnaire behavior.
6. Separate Khalil and BTG production brands.

## vX3 roadmap

1. CRM Automation MVP.
   - Provide a protected internal admin console for CRM routing and operational settings.
   - The administrator selects one active CRM: `amocrm`, `bitrix24`, or `off`.
   - CRM selection and diagnostics are never shown on the customer screen.
   - Implement and validate amoCRM first in the existing Test deployment.
   - After amoCRM acceptance, use the same Test admin console to configure and validate Bitrix24.
   - Never send one audit to both CRMs; switching providers changes the destination for future audits only.
   - Create or update the contact, company, and lead after a completed audit.
   - Attach audit artifacts and store industry, IT/IS maturity, priorities, and source application.
   - Create the first sales follow-up and presales tasks from confirmed P1 findings.
   - Prevent duplicate leads by normalized phone and email.
2. Collaborative Audit.
   - Store the questionnaire server-side behind a protected link.
   - Track author, timestamp, owner, and status for every section.
   - Preserve revision history and prevent accidental overwrites.
3. Live Audit QA.
   - Exercise Test, Khalil, and BTG after deployment using representative audit fixtures.
   - Verify draft loading, presentation generation, delivery, branding, and mobile rendering.
   - Notify maintainers only when a live check fails.
4. Quality and Sales Intelligence.
   - Validate evidence, recommendations, vendors, regulatory context, wording completeness, and roadmap consistency before delivery.
   - Add lead scoring, next-best action, call scenarios, objections, and presales tasks.

## CRM implementation principles

- All amoCRM and Bitrix24 development and validation are performed only in the existing Test deployment until explicit production approval.
- CRM failures must never block delivery of the customer presentation.
- Customer-facing screens must not expose CRM names, diagnostics, or internal sales data.
- The database access credential lives only in Streamlit secrets.
- CRM credentials entered in the admin console are stored in an encrypted server-side vault and are never returned to the browser after saving.
- The active provider is stored in the protected configuration database and defaults to `off`.
- Every delivery attempt receives an idempotency key and an internal status.
- CRM payloads contain only the data required for lead processing.

## Admin console

- The admin console is not linked from the customer questionnaire.
- Access requires an authenticated administrator session and an explicit email allowlist or a temporary Test-only admin password.
- The console can enable or disable CRM delivery, select an allowed provider, switch test mode, choose pipeline/status/responsible user, configure task deadlines, and run a connection test.
- The console selects whether the customer receives the management presentation, the expert Excel workbook, or both files.
- The console controls Telegram diagnostics, lead text, captions, and which generated files are sent.
- Provider credentials can be entered or replaced in the console, but saved values are masked and never displayed again.
- Persistent non-secret settings and delivery history are stored outside the Streamlit filesystem.
- A new CRM configuration remains inactive until its connection test succeeds and the administrator explicitly activates it.
- Switching CRM never resends historical audits automatically.
- Production CRM routing remains disabled until each provider passes its isolated acceptance test.

## CRM rollout sequence

1. `X3-dev.1`: protected admin console, persistent settings, normalized lead payload, and amoCRM connection diagnostics in Test.
2. `X3-dev.2`: amoCRM contact/company deduplication, lead creation, artifacts, and tasks.
3. `X3-dev.3`: amoCRM manual acceptance and failure handling.
4. `X3-dev.4`: Bitrix24 adapter using the same normalized payload and the same protected Test admin console.
5. `X3-dev.5`: switch Test to Bitrix24, run isolated Bitrix24 acceptance, then return the provider to `off` or amoCRM.
6. `X3-rc1`: regression and live checks for both saved provider configurations in Test; production remains unchanged.
