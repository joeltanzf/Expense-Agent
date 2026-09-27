# TNG Expense Agent

**Turn TNG eWallet PDF statements into organized Excel expense records on your Windows PC.**

TNG Expense Agent helps you track spending without copying every transaction into Excel yourself. Upload a supported statement, enter its password, and the app reads the date, description and amount. It saves expenses into the correct month's sheet in a separate Excel workbook for each year.

You can also enter expenses yourself, rename merchant descriptions, remove unwanted entries, and switch between years.

The installed app works locally. It does not require a Codex subscription, an API key, an AI service, Node.js, or a separate Python installation.

[Install the app](#install-and-start) · [User guide](#how-to-use-it) · [How it works](#how-it-works) · [Run from source](#run-from-source)

## Why I built this

I built TNG Expense Agent to make it easier to keep track of everyday spending. Instead of entering every line from a statement into a spreadsheet, I wanted a tool that could organize expenses by date, keep monthly totals, and let me choose the descriptions I use. I am sharing it here so others can use it and explore how it works.

## What the app does

| Feature | How it helps |
| --- | --- |
| Import PDF statements | Reads supported text-based TNG eWallet statements, including password-protected PDFs. |
| Track expenses | Includes supported purchases and outgoing transfers; excludes reloads, incoming money and GO+ movements. |
| Add entries manually | Save a date, description and price directly from the app. |
| Organize by year and month | Creates files such as `Expenses-2026.xlsx`, with sheets such as January and February. |
| Rename descriptions | Save rules such as `Example Market` → `Groceries` for existing and future expenses. |
| Avoid repeated imports | Skips previously imported PDFs and known transaction references in overlapping statements. |
| Review possible matches | Lets you decide whether a PDF expense matches an entry you previously typed yourself. |
| Delete expenses | Remove a selected manual or PDF expense after confirmation, with a database backup first. |
| Update totals | Keeps entries in date order and writes a SUM formula below each month's prices. |
| Keep backups | Retains up to 30 database snapshots and 30 previous workbook versions. |

## Install and start

The app is designed for **64-bit Windows 10 or Windows 11**, using Windows PowerShell 5.1 and .NET Framework 4.7.2 or newer.

1. Open this repository's **Releases** section and download **`TNG-Expense-Agent-Setup.exe`** from the latest release.
2. Run the installer and confirm the installation.
3. Open **TNG Expense Agent** from the Windows Start menu.
4. Choose **Upload PDF**, or enter your first expense manually.

The installer is approximately **33 MB** and includes Python, `pdfplumber`, `openpyxl` and their dependencies. Installation and normal app use do not download additional components or contact an online processing service.

You do not need Microsoft Excel to generate the files. To view them, use an application that supports `.xlsx` workbooks. Microsoft Excel is a separate product and is not included.

To run or modify the code yourself, follow [Run from source](#run-from-source) below.

## How to use it

### 1. Import a PDF

1. Select **Upload PDF**.
2. Choose an original TNG eWallet statement.
3. Enter the PDF password and select **Read PDF**. Leave the field blank to use a password already saved in Settings.
4. Review any possible matches with manual entries.
5. Check the import result, select the year, and choose **Open [year] Excel**.

A password entered in the upload dialog is used for that upload and is not saved. To remember it, enter it under **Settings → PDF password** and choose **Save settings**. Saved passwords use Windows encryption for your account. If Windows cannot save or unlock one, use the one-time upload field instead.

The parser supports the text-based TNG statement layout it was built for. Scanned PDFs, unfamiliar layouts and unknown transaction types stop for review without importing a partial statement. Use statements smaller than 30 MB with no more than 150 pages.

### 2. Add an expense yourself

In the Expenses tab:

1. Choose the **Date**.
2. Enter a **Description**.
3. Enter the **Price (RM)**, such as `12.50`.
4. Select **Save to Excel**.

The expense date determines its year and month. For example, an entry dated 15 January 2027 goes into the **January** sheet in **`Expenses-2027.xlsx`**, regardless of the year you were viewing.

### 3. Choose a year

Use **Previous year**, **Next year**, or the year dropdown. The expense list, saved-entry count and total change to the selected year. The month filter lets you narrow that year's list.

Select **Open [year] Excel** to open that year's workbook. A year becomes available when it has a monthly sheet, including when you add an expense dated in that year. Navigation buttons are disabled when there is no earlier or later saved year.

Example organization:

```text
Expenses-2026.xlsx
  January
  February
  ...

Expenses-2027.xlsx
  January
  ...
```

Each sheet has **Date** in A1, **Description** in B1, and **Price** in C1. Prices use RM formatting. A month could look like this:

| Date | Description | Price |
| --- | --- | ---: |
| 10/01/2026 | Groceries | RM 25.50 |
| 12/01/2026 | Lunch | RM 12.00 |
| | **Total:** | **RM 37.50** |

The total is an Excel SUM formula. A saved formula result also allows viewers that do not calculate formulas to display the last exported total.

### 4. Rename descriptions

Open **Description rules**, enter the merchant phrase and the replacement description, then select **Save rule**.

For example, a rule matching `Example Market` can write `Groceries` into column B. Choose **Contains** to match a phrase within a description, or **Exact** to match the whole description. Matching is case-insensitive. Exact rules take priority; otherwise, the first matching rule wins.

Rules apply to saved and future expenses. The original merchant description remains stored and visible in the app.

### 5. Delete an expense

1. Select its row in the expense list.
2. Select **Delete selected**.
3. Check the date, description and amount.
4. Choose **Delete expense**, or **Keep expense** to cancel.

This works for both manual and PDF entries. The app creates a database backup first, removes the active entry, and updates its yearly workbook and total. If the workbook is open, the deletion is saved and the Excel update waits until the file is available.

Deleted PDF references stay excluded from later overlapping imports. This also applies when a deleted manual entry had already been linked to a PDF. An unlinked manual entry has no known PDF reference, so a later statement can still introduce that transaction.

Backups and an older archived workbook may still contain deleted entries. Deletion is not secure erasure of every historical copy.

### 6. Handle duplicates

Repeated PDFs and known statement references are skipped. Different references remain separate even if the dates and amounts match. If a reference conflicts with saved transaction details, the import stops for review.

When a new PDF row has the same date and amount as an unlinked manual entry, choose:

- **Use selected entry:** keep the manual entry and link the PDF reference, counting the expense once.
- **Keep both:** save both because they are different purchases.
- **Cancel upload:** leave the whole PDF unimported.

### 7. Create new monthly sheets automatically

The app creates the current month on launch and checks for a new month hourly while open. It also tries to enable a Windows daily/sign-in task on its first normal launch. Check its status in **Settings**, or select **Enable automatic monthly sheets**.

If Windows does not allow the task, month creation still works when you open the app. A computer that is turned off cannot create a sheet until it starts again. The app catches up missed months after the latest saved non-future month.

## Where your data is stored

Press **Windows + R**, paste the following path, and press Enter:

```text
%LOCALAPPDATA%\TNGExpenseAgent\data
```

This folder contains:

| Item | Purpose |
| --- | --- |
| `transactions.sqlite3` | Saved expenses, description rules, import history and settings. |
| `Expenses-YYYY.xlsx` | The workbook for each year. |
| `backups` | Database snapshots and previous workbook versions. |

PDFs are read from the location you select. The app does not upload them or copy them into its installation folder. Transaction data and processing stay on your PC; any cloud synchronization you configure for your own folders is separate from the app.

Remembered passwords are encrypted for your Windows account. The database and Excel workbooks themselves are **not encrypted**, and database backups can contain older encrypted password values. Keep the data folder private and make your own backup on separate storage.

## Recovery and updates

- **Excel is open:** close the affected workbook and leave the app open so it can retry the update. Your changes remain saved in the database.
- **You edited Excel directly:** select that year, open Settings, and use **Back up and rebuild [year]**. This preserves a workbook backup and regenerates that year from saved app records. Direct Excel edits are not imported back into the app.
- **Update the standalone app:** close it and run the newer installer. Its separate data folder is retained.
- **Upgrade from the first standalone version:** yearly files are created from saved records. The old combined `Expenses.xlsx` remains unchanged as an archive and is no longer updated.
- **Move from an older folder-based app:** close both apps and Excel, back up the older `data` folder, and copy it into the new data folder only if the destination is empty. Do not overwrite a destination that already contains expenses; those databases need a deliberate merge.

To restore a database backup, close the app and Excel, make a copy of the current data folder, then replace `transactions.sqlite3` with the selected backup. Reopen the app and rebuild the affected years. Passwords saved under a different Windows account may need to be entered again.

## How it works

```text
TNG statement PDF + password
             |
             v
Read transaction columns with pdfplumber
             |
             v
Keep supported expenses, check references and review possible matches
             |
             v
Save transactions and description rules in SQLite
             |
             v
Create yearly Excel workbooks and monthly sheets with openpyxl
```

Manual entries and confirmed deletions update the same SQLite records and trigger workbook updates. Renaming rules change the exported descriptions while keeping original text available. New workbook files are validated before replacing the previous files.

| Component | Role |
| --- | --- |
| Python | Transaction processing and application logic. |
| pdfplumber | Reading text and transaction columns from supported PDFs. |
| SQLite | Local storage for expenses, rules and history. |
| openpyxl | Writing Excel workbooks, formatting and formulas. |
| Windows PowerShell and WPF | The desktop interface. |
| C# and .NET Framework | The launcher and offline installer. |

## Run from source

These steps are for people running or developing the source code. Installer users can skip them.

Install **64-bit Python 3.12**, download the source, and open PowerShell in the folder containing `requirements.txt` and `Launch Agent.ps1`.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Launch Agent.ps1"
```

Installing source dependencies requires an internet connection unless they are already available locally. Running the installed app does not. The launch command sets the script execution policy for that process only; organization policies may still prevent execution.

### Run the tests

Tests use fictional transactions and temporary storage.

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -v
```

### Build the installer

Use a complete standard Python 3.12 x64 installation with `requirements-dev.txt` installed, plus the Windows .NET Framework C# compiler. Run the build with that full Python installation:

```powershell
python build/build.py --out dist/release
```

The output folder must not already exist. The build produces the installer, its SHA-256 checksum and intermediate payload files. Installers and checksums are distributed through **GitHub Releases**, while the source code is kept in this repository.

## Testing

I checked the app manually before release. Automated testing of the version with year selection and expense deletion also produced these results:

- **32 backend tests passed**, with no skipped tests.
- **24 installation and integration checks passed**, including an upgrade from the previous installer using fictional data.
- UI checks covered year navigation, choosing the workbook to open, cancellation, and confirmed deletion.
- Microsoft Open XML validation found **zero errors** in the tested 2025 and 2026 workbooks.

The automated checks used fictional transactions in an isolated local Windows environment. They included workbook-format validation and checking that **Open [year] Excel** selects the correct file. They did not automate Microsoft Excel itself.

## Compatibility and limitations

The installer is unsigned. Windows or an organization's security policy may block it; follow the applicable policy and ask the maintainer for an approved build when needed. Unsupported PDF layouts and scanned statements require further development.

## Reporting an issue

Open an issue in this repository's **Issues** tab. Include what you clicked, the error message, your Windows version, and whether you used the installer or source version. For PDF problems, describe the statement layout or transaction type.

Do not post real statement passwords, financial PDFs, databases or personal workbooks in public issues. Use fictional or carefully redacted examples.

## Project files and notices

```text
app/                     Desktop interface, PDF processing and Excel export
scripts/                 Runtime lookup and monthly-check setup
tests/                   Tests using fictional data
build/                   Launcher, installer and build source
docs/                    Setup, user guide and architecture notes
Launch Agent.ps1         Source-version entry point
requirements.txt         App dependencies
requirements-dev.txt     Development and test dependencies
THIRD-PARTY-NOTICES.md    Dependency license information
README.md                This guide
```

Keep financial data, passwords, generated workbooks, bundled runtimes and build intermediates out of the source repository. Retain third-party license notices when distributing the installer. A component's license does not automatically apply to the project's own code.

This is an independent project and is not affiliated with or endorsed by Touch 'n Go or Microsoft.
