# Data Workbench — first implementation

A local Windows desktop application for importing data, reviewing it as a table, applying simple filters and cleaning presets, drawing charts, and exporting results. The app is a Python/PySide6 project rather than a prebuilt installer; follow the steps below to run it on a Windows PC.

## Start on Windows

Install Python 3.11 or 3.12 (64-bit) and select **Add Python to PATH** during installation. Place `data_workbench.py` and `requirements.txt` in the same folder. In PowerShell from that folder:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python data_workbench.py
```

If Python 3.11 is installed instead, replace `-3.12` with `-3.11`. To produce an optional standalone Windows executable, run this **on Windows** after installing the dependencies:

```powershell
pyinstaller --noconfirm --onefile --windowed --name DataWorkbench --collect-all matplotlib --collect-all pandas data_workbench.py
```

The result is `dist\DataWorkbench.exe`. This is a build command, not a tested installer. If a frozen build fails to import an optional file-format package, rebuild with `--collect-all <package>` for that package. `.doc` import still needs LibreOffice installed separately with `soffice` on PATH. Untrusted files should only be opened in an appropriately isolated environment.

## Workflow and uses

1. Click **Import file** and select a file. For Excel choose a worksheet; for SQLite/SQL choose a table. The imported data appears as a table. The workspace stores both the original import and the current processed version at `%LOCALAPPDATA%\DataWorkbench\workspace.sqlite3` (or in the home directory under `DataWorkbench` if `LOCALAPPDATA` is unset). Importing the same file again creates a separate saved dataset.
2. Enter a condition and click **Apply filter**, or select a cleaning preset and click **Apply preset**. Operations affect the current saved dataset and can be repeated. **Reset to import** restores its original rows.
3. Choose a chart type and columns, then click **Draw chart**. Save the chart as PNG if desired.
4. Click **Export** to save the currently processed dataset. Select a filename with the extension matching the chosen output type.

Typical uses include inspecting a spreadsheet, removing duplicate or incomplete rows, excluding extreme numeric observations before plotting, searching extracted document text, and moving cleaned tabular data between formats.

## Import and output formats

| Format | Import behavior | Export behavior |
| --- | --- | --- |
| CSV | Auto-detect common delimiters; UTF-8 then Windows-1252 fallback | UTF-8 with BOM |
| XLS / XLSX | Select one sheet | XLSX only |
| TXT | Try delimited table, otherwise one text line per row | Tab-separated text |
| PDF | Extract selectable text, one line per row with page number | Simple paginated table PDF (up to 12 columns) |
| DOC / DOCX | Extract paragraphs and table rows; DOC requires LibreOffice | DOCX only, as a table |
| JSON | Flatten list of objects, a first list-of-objects property, or a single object | Array of row objects |
| XML | Flatten immediate child elements and attributes into rows | Row-oriented XML |
| SQL / DB | SQLite database files (`.db`, `.sqlite`, `.sqlite3`) or trusted SQLite-compatible `.sql` dump; select a table | SQLite `.db` or SQLite `.sql` dump |
| Chart | Not imported | PNG |

## Filtering and charts

Filter syntax is `[column] operator value`, optionally joined with `AND` (case insensitive). Available operators are `=`, `!=`, `>`, `>=`, `<`, `<=`, `~` (case-insensitive literal substring). Examples:

```text
[age] >= 18 AND [region] = West
[city] ~ York
[status] != null
```

`null` checks missing values with `=` or `!=`. Text equality ignores case; numeric comparisons coerce column values to numbers and exclude nonnumeric values. Filter conditions are applied sequentially to the *current* data. Values are unquoted literal strings; there is no OR, parentheses, natural-language filter, or arbitrary Python/SQL execution.

Presets: remove exact duplicate rows; drop rows with any missing cell; remove any row whose numeric value lies outside 1.5 times the IQR from the first/third quartile for any varying numeric column. The IQR rule is a heuristic, not evidence that a row is erroneous. Charts: histogram, top-30 category bar chart, line chart, scatter chart. The grid previews the first 500 rows; charts and exports use all current rows.

## Limitations and cautions

- No external AI service is connected. Specialized natural-language prompts are a future extension; the first version uses the explicit filter grammar above and requires no network connection.
- No OCR: scanned/image-only PDFs may produce no rows. PDF/DOC text import is not faithful reconstruction of document tables, layout, formulas, embedded charts, or images. DOC-to-DOCX conversion can lose formatting.
- SQL support means **SQLite**, not arbitrary database servers (SQL Server, MySQL, PostgreSQL) and not their non-SQLite dump syntax. SQL dumps execute in an in-memory SQLite database after a trust confirmation; attaching external databases and PRAGMA/VACUUM commands are denied, but only run dumps you trust. Existing database files are opened read-only.
- XML import is shallow and may combine repeated child tags. JSON nested lists can remain in cells; exports may stringify complex objects. Inputs are imported into memory; very large files and workspaces can be slow or consume substantial RAM/disk. No streaming, pagination beyond the 500-row grid preview, schema editor, or undo history; **Reset to import** discards all processing for the selected dataset.
- Type inference and round trips can change data types or formatting, especially dates, leading-zero identifiers, decimals, and long numeric IDs. Verify critical data against its source. CSV formula cells can be interpreted by spreadsheet programs when opened there; do not treat exports as sanitized input. Document/PDF exports suit small, narrow tables and can truncate displayed cell text.
- All imported and processed data is persisted locally as compressed blobs in the workspace SQLite file, including potentially sensitive content. No encryption, user accounts, sharing, or cloud sync is provided. Secure or delete the workspace as needed.
- This is an initial implementation, not a production-hardened ETL system. The script and packaging instructions are supplied for local execution; a signed Windows installer and automated end-to-end validation are not provided.
