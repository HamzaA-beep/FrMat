# FrMat

FrMat is a privacy-first local desktop data tool for importing messy files, reviewing them as tables, cleaning the data, generating quick charts, and exporting the final result.

It is designed for small-scale analysis without requiring a cloud backend, external AI service, or account setup. Everything happens locally on the machine where the data lives.

## Why FrMat

FrMat helps with everyday data tasks such as:

- reviewing CSV, Excel, PDF, DOCX, JSON, XML, SQL, and SQLite data
- filtering rows by condition
- removing duplicates or incomplete records
- dropping outlier values before analysis
- drawing simple visual summaries
- exporting clean results for downstream use

## Key features

- Import structured and semi-structured file types
- View imported data in a tabular interface
- Apply filters with a plain-language style query format
- Use built-in cleanup presets
- Remove outlier rows with IQR-based logic
- Generate histogram, bar, line, and scatter charts
- Export to CSV, XLSX, JSON, XML, SQLite, DOCX, and PDF
- Keep all imported and processed data stored locally

## Example filter syntax

```text
[age] >= 18 AND [region] = West
[city] ~ York
[status] != null
```

## Requirements

- Python 3.11 or 3.12 (64-bit recommended)
- Windows is the primary supported platform
- Local access to package installation via `pip`

## Quick start on Windows

From PowerShell in the project folder:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python data_workbench.py
```

If Python 3.11 is installed instead, replace `-3.12` with `-3.11`.

## Optional standalone build

To build a one-file Windows executable with PyInstaller:

```powershell
py -3.12 -m pip install --upgrade pip pyinstaller
pyinstaller --noconfirm --onefile --windowed --name FrMat --collect-all matplotlib --collect-all pandas data_workbench.py
```

The output is created in the `dist` folder.

## Supported formats

| Format | Import | Export |
| --- | --- | --- |
| CSV | Yes | Yes |
| XLS / XLSX | Yes | Yes |
| TXT | Yes | Yes |
| PDF | Yes | Yes |
| DOC / DOCX | Yes | Yes |
| JSON | Yes | Yes |
| XML | Yes | Yes |
| SQL / SQLite | Yes | Yes |
| PNG chart | No | Yes |

## Project structure

```text
FrMat/
├── LICENSE
├── PACKAGING.md
├── README.md
├── requirements.txt
├── data_workbench.py
├── installer/
│   └── Install-From-Source.ps1
├── scripts/
│   └── build_windows.ps1
├── docs/
│   ├── overview.md
│   ├── components.md
│   ├── installation-guide.md
│   ├── maintenance.md
│   └── troubleshooting.md
└── .gitignore
```

## Typical workflow

1. Import a file such as CSV, Excel, PDF, or SQLite.
2. Review the table and inspect column names.
3. Apply a filter or cleaning preset.
4. Draw a quick chart for visual checks.
5. Export the cleaned file in the desired format.

## Limitations

- This is a local desktop utility, not a full production ETL platform.
- PDF and DOC extraction depends on readable text in the source file.
- SQL support is limited to SQLite-compatible inputs.
- Very large or highly complex files should be validated against their source data.
- The app stores imported and processed data locally in a SQLite workspace database.

## License

This project is licensed under the MIT License. See `LICENSE` for details.
