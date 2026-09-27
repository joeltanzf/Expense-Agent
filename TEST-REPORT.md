# Year selection and expense deletion: test report

Date: 2026-09-26

## Delivered changes

1. Select a manual or imported PDF expense and click **Delete selected**. A confirmation shows its date, description and amount. A consistent database backup is made before deletion. The expense disappears from active records and its yearly workbook; the total updates. Known PDF references remain excluded from overlapping imports, including references previously linked to a manual entry. A new manual submission may deliberately add an expense again.
2. Use **Previous year**, **Next year**, or the year selector. Counts, totals and the expense list reflect that year. **Open [year] Excel** opens `Expenses-YYYY.xlsx`, containing only that year's month sheets. Dates from manual entries and PDFs choose their year automatically. Rebuilding a selected year preserves outside edits in other years.

The old combined `Expenses.xlsx` is retained unchanged during upgrade and is no longer updated. The app's saved database is the source for the new yearly workbooks. Old direct Excel edits are not imported automatically. Existing standalone data stays in the same per-user folder; older folder-based installations still follow the migration guide.

## Results

**32 backend tests passed, zero skipped.** They ran against the installed app with its bundled Python. Coverage includes the existing PDF/password/manual/rule/duplicate/backup cases plus deletion of manual, PDF and linked records; preventing deleted PDF references from returning; safe cancellation and missing-selection handling; backup failure preventing deletion; stale manual retry rejection; locked workbook recovery after deletion; separate year totals/month names; independent year locking; protection of existing unowned files; selected-year rebuild protection; and older database migration.

**24 installation and integration checks passed.** The upgrade test installed the actual previous installer into an isolated folder, created fictional entries in two years, then installed this update. It verified that the installer left the database unchanged, that the app retained the entries and generated yearly files, and that the old combined workbook bytes remained unchanged. A second installation also preserved the test database.

Actual WPF button tests exercised Previous/Next year, navigation boundaries, year filtering, selecting the Open Excel destination, cancelled deletion, and confirmed deletion through the real background worker. The Open Excel destination was intercepted for verification; this did not open Microsoft Excel. The deletion-confirmation dialog itself passed confirm/cancel interaction tests. Earlier password and possible-duplicate dialogs were also exercised. The rendered updated interface was visually inspected.

**Microsoft Open XML validation: zero errors** in both the 2025 and 2026 workbooks produced by the installed app. SUM formulas and cached totals remain covered by the backend tests.

## Remaining acceptance checks

These were isolated local Windows tests, not a clean Windows VM. Native Microsoft Excel opening/recalculation, normal-account password remembering, and actual Start menu/scheduled-task registration still need normal desktop acceptance checks, as noted for the earlier release. The monthly-check backend route passed. The installer remains unsigned.

## Artifact and privacy

Installer: `TNG-Expense-Agent-Setup.exe` (32.9 MiB).

The same free Python/openpyxl/pdfplumber components are bundled; no extra user download, Codex subscription, Node.js or API key is required. The runtime provenance and third-party licensing are unchanged from the first candidate: the included CPython came from the build environment and has an OpenAI signing certificate; its license and public dependency notices are retained. The installer itself is unsigned.

No real expenses, statements, passwords or workbooks were used or included. The earlier standalone source (23 files), original GitHub-ready source and earlier prototype (37 files) were checked unchanged. No personal data was migrated, deleted or overwritten during testing. Nothing was published to GitHub.

Backups and the old combined workbook may retain deleted records. Deletion affects the current app records and yearly workbooks; it is not secure erasure. A manual entry with no known PDF reference cannot automatically suppress a future PDF transaction.

SHA-256: `e24e8308d31f8bce3249b8d80d9cf0a2b21964e56a9aaa2c2afa5af817e9b555`
