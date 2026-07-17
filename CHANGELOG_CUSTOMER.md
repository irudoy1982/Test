# Customer-facing changelog

This changelog is intended for production release notes that can be shown to customers.
Keep this file focused only on customer-visible questionnaire and report changes.

## v12

Status: in development in Test.

Planned customer-visible improvement:
- Added a concise branded PowerPoint presentation with the audit summary, infrastructure profile, key risks, IT and security recommendations, 90-day roadmap, and management decisions.
- Added separate IT and information security maturity indicators to the presentation cover.
- Added solution classes and relevant manufacturers to presentation recommendations.
- Added a clearer final action slide with company contacts.
- Reworked the infrastructure slide to explain the operational priorities created by the customer's environment.
- Added a recommended action to every key risk in the presentation.
- Unified the recommendation slides, removed semantic duplicates, and ensured that actions are shown as complete statements.
- Improved presales wording so presentation risks, impacts, and actions are concise and complete.
- Limited recommended manufacturers to the maintained solution portfolio.
- Improved the 90-day plan so later phases continue the identified priorities instead of showing generic actions.
- Simplified customer delivery to one management presentation.
- Moved the company profile near the conclusion so the presentation opens with the customer's situation and decisions.
- Made the company's delivered IT project track record more prominent.
- Refined company founding facts and improved the audit conclusion download experience.
- Polished the final audit conclusion action for a clearer, more focused download.
- Centered the final download action and improved its responsive behavior.
- Centered and refined the presentation generation action so the primary next step is clear and visually consistent.
- Added sector-aware Kazakhstan regulatory context and clearer separation between legal obligations and recommended standards.
- Expanded the sector list with banking, insurance, healthcare, public and quasi-public sectors, critical infrastructure, telecom, utilities, transport, and industrial environments.
- Rebuilt the presentation as a 13-slide decision narrative with confirmed strengths, a consistent severity palette, regulatory context, evidence-based recommendations, and measurable outcomes.

## v12.0.0

Status: stable production release.

Customer-visible improvements:
- Clearer report structure for the customer-facing audit workbook.
- More practical recommendations tied to the scale of the customer's IT infrastructure.
- Better separation between IT maturity and information security maturity.
- Improved roadmap wording and prioritization.
- Cleaner generation flow so users can see that report creation has started.

Implemented:
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
