# Customer-facing changelog

This changelog is intended for production release notes that can be shown to customers.
Keep this file focused only on customer-visible questionnaire and report changes.

## v12.0-dev

Status: in development in Test only.

Planned customer-visible improvements:
- Clearer report structure for the customer-facing audit workbook.
- More practical recommendations tied to the scale of the customer's IT infrastructure.
- Better separation between IT maturity and information security maturity.
- Improved roadmap wording and prioritization.
- Cleaner generation flow so users can see that report creation has started.

Implemented in Test:
- Added IT context to the audit workbook passport: endpoints, servers, virtualization, storage, public services, business systems, and operational focus.
- Improved recommendation flow so audit outputs stay aligned with the same findings.
- Reworked network findings so routing protocols are not treated as proof of missing segmentation; unconfirmed architecture is now marked for validation.
- Improved matching between identified risks and solution categories in the customer report.
- Replaced subjective company-size labels with objective infrastructure facts from the questionnaire.

## v11.01-prod

Status: stable production baseline.

Customer-visible changes:
- Improved questionnaire flow and draft handling.
- Added clearer audit navigation and section completion states.
- Improved XLSX report layout and risk presentation.
- Added portfolio-based vendor matching for customer recommendations.
- Improved report generation flow and duplicate-click protection.
