# UAR & SoD Review Tool — Oracle ERP

Single-file, browser-only tool for User Access Review and Segregation of Duties analysis on Oracle ERP (EBS) access extracts. Open `uar-tool.html` in a browser — no install, no server, all processing stays on your machine.

## Usage

1. Open `uar-tool.html` (needs internet once per open, for the SheetJS CDN and fonts).
2. Upload files into the four slots:
   - **UAR file** — sheets with Overall / Joiner / Exit employee lists
   - **Oracle responsibility extract** — user-responsibility assignments (multiple sheets merge automatically); SQL for the extract is in the repo history of your data file
   - **Name/email mapping** — PERSON_ID / LAST_NAME / EMAIL_ADDRESS / FULL_NAME
   - **Custom SoD matrix (optional)** — `Category A | Category B | Risk | Description` (+ optional classification sheet). If empty, the built-in Oracle EBS matrix is used.
3. Click **Run review**.

## Checks

- **SoD conflicts** — responsibility-level, 18 built-in Oracle EBS conflict rules across 18 access categories (AP/GL/PO/AR/OM/INV/HR/…), generic-account and super-user flags. Inquiry/read-only responsibilities excluded. Rules editable in the UI.
- **Leaver check** — exited employees with responsibilities still active after exit date (matched via email local-part → Oracle username).
- **Joiner check** — new joiners' granted access for appropriateness review.
- **Full UAR dump** — every user-responsibility row with employee match and status.
- **Verdict column** — mark each finding True/False Positive; flows into exports.

## Outputs

- Dashboard with severity/rule/module charts
- Excel export (one sheet per check)
- Word report (audit working paper)

## Data privacy

Excel data files are git-ignored — never commit access extracts or employee lists.
