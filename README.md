# FrMat

FrMat is a local desktop data workbench for importing messy files, reviewing them in a table, cleaning and filtering them, generating quick charts, and exporting the final result.

It runs directly on your desktop machine and is designed for practical local analysis without a cloud backend or remote service.

## Runs on your desktop

FrMat is a desktop-first tool. It is intended to be used on a local Windows workstation, where the data stays on the machine and the app runs from the user’s own environment.

## Minimum hardware requirements

- 64-bit CPU
- 4 GB RAM minimum
- 8 GB RAM recommended for smoother performance
- 1 GB free storage space minimum
- No dedicated GPU required for normal use

## Minimum software requirements

- Windows 10 or Windows 11 (64-bit)
- Python 3.11 or 3.12
- pip access for installing project dependencies
- Optional: PyInstaller for building a standalone Windows executable
- Optional: LibreOffice `soffice` on PATH for `.doc` import conversion

## Why FrMat

FrMat helps with common local data tasks such as:

- reviewing CSV, Excel, PDF, DOCX, JSON, XML, SQL, and SQLite data
- filtering rows with simple conditions
- removing duplicate or incomplete records
- dropping outlier values before analysis
- drawing quick charts for review
- exporting cleaned results for downstream use

## Key features

- Import and inspect common file formats
- View data in a table interface
- Apply filters using a simple expression format
- Use built-in cleanup presets
- Remove outlier rows with IQR-based logic
- Generate histogram, bar, line, and scatter charts
- Export to CSV, XLSX, JSON, XML, SQLite, DOCX, and PDF
- Keep the workflow local to the desktop environment

## Example filter syntax

```text
[age] >= 18 AND [region] = West
[city] ~ York
[status] != null
```

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
2. Review the table and inspect the loaded columns.
3. Apply a filter or cleanup preset.
4. Draw a quick chart for visual checks.
5. Export the cleaned result in the desired format.

## Limitations

- This is a local desktop utility, not a full production ETL platform.
- PDF and DOC extraction depends on readable text in the source file.
- SQL support is limited to SQLite-compatible inputs.
- Large or highly complex files should be validated against their source data.
- The app stores imported and processed data locally in a SQLite workspace database.

## License

This project is licensed under the MIT License. See `LICENSE` for details.

The MIT License is a good choice for a public GitHub project if you want it to be broadly reusable and easy for others to build on. It allows modification, redistribution, and commercial use while requiring attribution and preserving the copyright notice.
