# Architecture

The WPF interface runs in built-in Windows PowerShell. A small .NET Framework launcher starts it without a console. Python parses text-based TNG PDFs with pdfplumber, stores records and rules in SQLite, and writes ordinary XLSX workbooks using openpyxl. No web service, AI model or API is called.

The Python process accepts one JSON request on standard input and returns JSON on standard output. One-time passwords travel through this pipe rather than the command line. A Windows file lock serializes backend operations. Windows DPAPI protects remembered passwords for the current user. No plaintext fallback is used.

Imports validate every row before committing. Statement-reference identity handles overlap; pdf_matches links manually reviewed matches without double counting. Decimal comparison treats equivalent amount spellings consistently. Text is checked before database changes; formula-like descriptions are literal strings in Excel.

Workbooks are written to a temporary file and validated before an atomic replacement. Existing external edits are detected by SHA-256. Old workbook contents and consistent SQLite snapshots are backed up. Data storage is independent of the installation directory.

The `year_exports` table tracks fingerprints, payload hashes and pending status separately for every year. Each yearly workbook contains that year only, with plain month names. An unchanged year is not regenerated; a locked or externally edited year cannot block other years. Explicit rebuild applies to the selected year. The prior combined workbook is retained as an untouched legacy archive.

Deletion takes a consistent database snapshot first, then removes the selected active transaction in a database transaction. `deleted_transactions` retains only identifiers and deletion timestamps. Known linked PDF identifiers are recorded there before the match metadata is removed, so overlapping statements cannot restore a deleted expense. A stale retry of a deleted manual submission is rejected; a new manual submission has a new identifier.

## Build

Use a standard complete Python 3.12 x64 installation with `requirements-dev.txt` installed, and Windows .NET Framework's C# compiler. Run `python build/build.py --out <new-output-directory>`. The build copies only the standard Python runtime and distributions required by pdfplumber/openpyxl, retaining their license metadata. It compiles the launcher and installer with the local C# compiler. It does not require PyInstaller, NSIS, Node or a proprietary spreadsheet package.

The output contains a single installer, a SHA-256 checksum, and intermediate payload files. Distribute the installer and checksum as release assets; keep the source folder as a normal repository. Never commit runtime binaries, data, workbooks, PDFs, passwords or build intermediates into the source repository.

The installer extracts a versioned payload and changes the active version only after extraction. Its internal `--test-install <empty-path>` mode installs without shortcuts, task registration or app launch for isolated testing. The installed launcher provides `--self-test`, `--ui-self-test` and `--ensure-month` for diagnostics. Normal launches set a stable installation path for the scheduled task across updates.
