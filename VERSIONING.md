# Versioning

## Release lanes

- `v12.0.0`: current production label from the previous versioning scheme.
- `v12.2-dev`: current development lane in Test with the refined presentation.
- `v12-rcN`: release candidate after Test validation, bug fixes only.
- `v12`: stable release after production approval.

## Rules

- Work on new features only in `Test`.
- Do not modify `KhalilAudit` or `BTGAudit` during experiments.
- Promote to production only after the user explicitly approves it.
- Customer-facing changes go to `CHANGELOG_CUSTOMER.md`.
- Internal changes can be tracked in commits and technical notes, but must not be shown to customers.
- Minor releases use one digit after the point: `v12.1`, `v12.2`, `v12.3`.
- Development and release-candidate suffixes are retained: `v12.1-dev`, `v12.1-rc1`.
- Major version `13` is written as `X3`: `vX3-dev`, `vX3-rc1`, `vX3`.
- The digit `3` in minor versions is not replaced: `v12.3` is valid.

## v12 direction

1. Version label and customer changelog.
2. AI-first sales playbook recommendations, with deterministic fallback only when AI is unavailable.
3. Automated smoke checks for draft load, report generation, workbook sheets, vendor/distributor matching, and basic Telegram delivery status.
4. Customer report structure improvements.
5. Presentation export after brand decks are provided.
