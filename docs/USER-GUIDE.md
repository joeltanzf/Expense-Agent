# User guide

## Upload a statement

Choose Upload PDF, select an original text-based TNG eWallet statement, and enter its password. Leave the password field blank to use the password saved in Settings. Upload passwords are not saved automatically.

The app uses the wallet, section and transaction reference to detect duplicates. Two genuinely different references remain separate even when the date and amount are the same. Unsupported layouts, scanned PDFs and unknown transaction types stop for review without importing a partial statement.

If a PDF row matches the date and amount of an unlinked manual entry, review the descriptions. **Use selected entry** keeps your manual entry and links the PDF reference so it is counted once. **Keep both** saves both purchases. **Cancel upload** saves none of that PDF. The same manual entry can be linked only once.

## Add an expense yourself

Choose Date, enter Description and Price (for example 12.50), then choose Save to Excel. The date determines the sheet. Retry protection prevents an accidental repeated save of the same submitted entry.

## Delete an expense

Select its row in Expenses and choose **Delete selected**. Check the date, description and amount, then choose **Delete expense**. **Keep expense** cancels. This works for manual entries and PDF imports. The app backs up the database before deletion and updates the yearly workbook and its total. If Excel is open, deletion is saved immediately and the workbook updates after you close it.

Deleted PDF references remain excluded from future overlapping statements. Deleting a manual entry that was linked to a PDF also excludes that PDF reference. A manual entry that has never been linked has no PDF reference to exclude; a later PDF can still introduce that transaction. Entering a new manual expense is always a separate action.

Backups may still contain deleted entries. This feature removes an expense from the active records and yearly workbooks; it is not a secure erasure feature.

## Rename merchants

In Description rules, enter a merchant phrase and the replacement name. Choose Contains or Exact. Exact matches take priority, then the first matching rule wins. Rules change column B for saved and future entries. Original descriptions remain stored.

## Years, monthly sheets and totals

Use **Previous year**, **Next year**, or the year selector to move between saved years. The list, count and total show that year. The month filter applies within it. **Open [year] Excel** opens `Expenses-YYYY.xlsx` for the selected year. A date entered manually or read from a PDF determines its year and month automatically. A year appears when it has a monthly sheet; empty years left after deletion remain available.

Each year has its own workbook, and sheets inside it use ordinary month names. 

Each sheet has Date at A1, Description at B1 and Price at C1. Entries start on row 2 and are sorted by date. Total appears below the entries and uses SUM. A saved formula result also supports viewers that do not recalculate. Excel recalculates when you edit prices, but make permanent changes through the app: direct Excel edits are not imported back into its database.

On first normal launch, the app tries to enable a Windows task for a daily and sign-in check. Settings shows its status. If Windows blocks task registration, use Enable automatic monthly sheets; the app still creates the current month on launch and checks hourly while open. A computer that is off cannot create a sheet until it starts again.

## Backups and recovery

If Excel is open, close the workbook and leave the app open. It retries pending updates. If the workbook was edited outside the app, select that year and choose **Back up and rebuild [year]** in Settings. This backs up and rebuilds only that year using saved transactions and rules. Other years with outside edits stay protected.

The private data folder is `%LOCALAPPDATA%\TNGExpenseAgent\data`. Copy the whole folder to your own secure backup location while the app is closed. Automatic backups retain 30 database snapshots and 30 prior workbook versions. They are on the same computer and do not protect against losing that computer.

To restore, close the app and the workbook, keep a copy of the current data folder, and replace `transactions.sqlite3` with the desired database backup. Reopen, select each affected year, and choose Back up and rebuild [year]. Passwords saved under another Windows account may need to be entered again.
