# UAR & SoD Review Tool — Oracle ERP

Single-file, browser-only, offline tool for User Access Review and Segregation of Duties analysis on Oracle ERP (EBS) access extracts. Open `uar-tool.html` in a browser — no install, no server, all processing stays on your machine.

## Usage

> **Current scope: SoD only.** The UAR checks (HR file upload, Leavers, Joiners) are hidden. Set `const SHOW_UAR = true;` in `uar-tool.html` to bring them back.


1. Open `uar-tool.html` in Chrome or Edge. Works fully offline: the SheetJS library (0.20.3, Apache-2.0) and the Outfit font are embedded, and the page makes no network requests.
2. Upload files into the slots:
   - **UAR file** — sheets with Overall / Joiner / Exit employee lists
   - **Oracle responsibility extract** — `User Name | Responsibility Name | Start Date | End Date | Responsibility Key | Application Short Name`, optional `EMAIL_ADDRESS` (used to link users to people; `EMPLOYEE_ID` is ignored). Add `Function Code | Function Name` columns (one row per function, query Q1) to switch on Level 3 — same upload, the tool removes the repeated responsibility rows. Multiple sheets merge automatically; SQL for all extracts is in `sql/query_stack.sql`
   - **Name/email mapping** — PERSON_ID / LAST_NAME / EMAIL_ADDRESS / FULL_NAME
   - **Custom SoD matrix (optional)** — `Category A | Category B | Risk | Description` (+ optional classification sheet). If empty, the built-in Oracle EBS matrix is used.
3. Click **Run review**.

## Checks

- **SoD conflicts** — responsibility-level, 62 built-in Oracle EBS conflict rules across 31 access categories (AP/GL/PO/AR/OM/INV/HR/payroll/bank/projects/…), including all 46 conflict pairs from the City of Atlanta Oracle SoD audit (2016), generic-account and super-user flags. Inquiry/read-only responsibilities excluded. Rules editable in the UI.
- **Level 3 (function-level) SoD** — when the extract has function columns, every responsibility-level finding is marked Confirmed / Not confirmed against the functions actually reachable through the menus, and conflicts visible only at function level (e.g. invoice + payment inside one responsibility) are added as "L3 only". Function → category map is editable in the Rules tab.
- **Leaver check** — exited employees with responsibilities still active after exit date (person → Oracle user by exact EMAIL_ADDRESS match, else username = email local-part; all of a person's accounts are checked).
- **Joiner check** — new joiners' granted access for appropriateness review.
- **Full UAR dump** — every user-responsibility row with employee match and status.
- **Verdict column** — mark each finding True/False Positive; flows into exports.

## Outputs

- Dashboard with severity/rule/module charts
- Excel export (one sheet per check)
- Word report (audit working paper)

## Data privacy

Excel data files are git-ignored — never commit access extracts or employee lists.
