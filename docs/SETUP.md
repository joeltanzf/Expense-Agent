# Setup and updates

## Installer

Run TNG-Expense-Agent-Setup.exe. It installs for the current Windows user without administrator access, under `%LOCALAPPDATA%\Programs\TNGExpenseAgent`. It creates a Start menu shortcut. All application dependencies are included; installation does not contact any server.

Expense data lives separately at `%LOCALAPPDATA%\TNGExpenseAgent\data`. Re-running the installer switches to its bundled version without deleting expense data. Old application versions are retained for recovery.

Upgrading the first standalone installer uses the same saved transactions, description rules, password settings and import history. On launch, the new app creates `Expenses-2025.xlsx`, `Expenses-2026.xlsx`, etc. from saved records. The earlier combined `Expenses.xlsx` is preserved unchanged and stops updating. Edits made directly in that older file are not imported into the database. Use the year buttons in the new app to open the current files.

The current installer is an unsigned release candidate. Organization policies or Windows reputation checks may prevent it from running. Do not disable organization security controls; ask the maintainer for an approved, signed release.

## Move data from the older folder-based app

Close both apps and Excel. Make an untouched backup of the older app's `data` folder. If the new data folder is empty, copy the older folder's contents into `%LOCALAPPDATA%\TNGExpenseAgent\data`. If the new folder already has transactions, do not overwrite it: the two databases need a deliberate merge. Reopen the new app and check the yearly files. Select a year and use its rebuild button if needed. Saved passwords remain tied to the original Windows user. Keep the old copy until you verify the totals.

## Remove the app

Close it. In Windows Task Scheduler, disable/remove only the TNG Expense Agent task whose action points to this installation. Delete the installation folder and its Start menu shortcut. Keep `%LOCALAPPDATA%\TNGExpenseAgent\data` if you want to preserve your expenses. The installer deliberately does not delete financial data.

## Source setup

Use Python 3.12 x64, create `.venv` in the source root and install `requirements.txt`. Run Launch Agent.ps1 with Windows PowerShell 5.1. The source launcher looks only for its own `.venv` or packaged runtime.

For isolated tests, `TNG_AGENT_DATA` and `TNG_AGENT_WORK` override storage locations. These are testing options, not a requirement for normal use.
