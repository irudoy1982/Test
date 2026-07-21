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
   - The CRM provider is selected internally per deployment: `amocrm`, `bitrix24`, or `off`.
   - CRM selection and diagnostics are never shown on the customer screen.
   - The existing Test deployment is reserved for amoCRM development and validation.
   - Bitrix24 is implemented only after amoCRM is accepted and receives a separate Test Bitrix deployment with isolated secrets and test records.
   - Never send one audit to both CRMs unless this is explicitly enabled in a future release.
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

- All amoCRM development and validation are performed only in the existing Test deployment until explicit production approval.
- Bitrix24 validation uses a separate Test Bitrix Streamlit application, not a provider switch in the existing Test deployment.
- CRM failures must never block delivery of the customer presentation.
- Customer-facing screens must not expose CRM names, diagnostics, or internal sales data.
- Credentials and account identifiers live only in Streamlit secrets.
- The active provider is controlled by the `CRM_PROVIDER` secret and defaults to `off`.
- Every delivery attempt receives an idempotency key and an internal status.
- CRM payloads contain only the data required for lead processing.

## CRM rollout sequence

1. `X3-dev.1`: normalized lead payload, provider switch, and amoCRM connection diagnostics in Test.
2. `X3-dev.2`: amoCRM contact/company deduplication, lead creation, artifacts, and tasks.
3. `X3-dev.3`: amoCRM manual acceptance and failure handling.
4. `X3-dev.4`: Bitrix24 adapter using the same normalized payload and a separate Test Bitrix deployment.
5. `X3-dev.5`: Bitrix24 acceptance without changing or contaminating the amoCRM Test environment.
6. `X3-rc1`: regression and live checks in both isolated test deployments; production remains unchanged.
