# Versioning

## Release lanes

- `v12.0.0`: current stable production release in KhalilAudit and BTGAudit.
- `v12.1.0-dev`: next development lane in Test only.
- `v12.1.0-rcN`: release candidate after Test validation, bug fixes only.

## Rules

- Work on new features only in `Test`.
- Do not modify `KhalilAudit` or `BTGAudit` during experiments.
- Promote to production only after the user explicitly approves it.
- Customer-facing changes go to `CHANGELOG_CUSTOMER.md`.
- Internal changes can be tracked in commits and technical notes, but must not be shown to customers.

## v12 direction

1. Version label and customer changelog.
2. AI-first sales playbook recommendations, with deterministic fallback only when AI is unavailable.
3. Automated smoke checks for draft load, report generation, workbook sheets, vendor/distributor matching, and basic Telegram delivery status.
4. Customer report structure improvements.
5. Presentation export after brand decks are provided.
