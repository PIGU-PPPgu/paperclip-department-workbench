# Changelog

## 1.1.1 — 2026-10-13

- Show a compact availability summary and each Paperclip agent's live status, manual or scheduled run mode, and blocked tasks on the workbench overview.
- Link agent and blocked task rows to their native Paperclip pages without adding agent write permissions to the plugin.

## 1.0.9 — 2026-10-13

- Match Paperclip's native typography, spacing, colors, borders, and buttons across the workbench.
- Show creation and editing forms in a native-style dialog.
- Fix workbench links in the dashboard widget and project tab so they retain the organization prefix.

## 1.0.5 — 2026-10-13

- Add native Paperclip workbench views for calendar, journal, OKR, milestones, knowledge, review, and monthly reports.
- Store user-created records in a plugin-owned PostgreSQL namespace with company scope, revision checks, and author metadata.
- Add creation and editing for meetings, development records, training, Goals, KRs, milestones, documents, and versions.
- Require human company membership for writes; keep pending records out of confirmed monthly totals.
- Preserve KR check-in history and milestone completion dates for monthly reporting.
- Remove the synthetic dataset and demo UI.
