# FrMat

FrMat is a local desktop application for importing data, reviewing it as a table, applying simple filters and cleanup steps, generating charts, and exporting the results.

This project is a lightweight data analysis tool that keeps all work on the user’s machine and does not require any cloud service or external AI backend.

## Features

- Import data from CSV, Excel, TXT, PDF, DOC/DOCX, JSON, XML, SQL, and SQLite files
- Review imported data in a table interface
- Apply row filters with a simple expression language
- Use built-in cleanup presets such as deduplication and missing-value removal
- Remove outlier rows based on IQR logic
- Draw basic charts from numeric and categorical data
- Export processed data to CSV, Excel, JSON, XML, SQLite, DOCX, or PDF formats
- Keep imported and processed data local to the machine

## Example filter syntax

```text
[age] >= 18 AND [region] = West
[city] ~ York
[status] != null
```

The filter format uses bracketed column names and operators such as `=`, `!=`, `>`, `>=`, `<`, `<=`, and `~`.

## Requirements

- Python 3.11 or 3.12 (64-bit recommended)
- Windows is the primary target for this project
- Local access to package installation via `pip`

## Quick start on Windows

From a PowerShell terminal in the project folder:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python data_workbench.py
```

If Python 3.11 is installed instead, replace `-3.12` with `-3.11`.

## Optional Windows build

To build a one-file Windows executable with PyInstaller:

```powershell
py -3.12 -m pip install --upgrade pip pyinstaller
pyinstaller --noconfirm --onefile --windowed --name FrMat --collect-all matplotlib --collect-all pandas data_workbench.py
```

The build output is placed in the `dist` folder.

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
└── .gitignore
```

## Typical workflow

1. Import a file such as CSV, Excel, PDF, or SQLite.
2. Review the table and check the loaded columns.
3. Apply a filter or preset to clean the dataset.
4. Draw a chart for quick visual checks.
5. Export the cleaned result in the desired format.

## Supported import/export formats

| Type | Import | Export |
| --- | --- | --- |
| CSV | Yes | Yes |
| XLS / XLSX | Yes | Yes |
| TXT | Yes | Yes |
| PDF | Yes | Yes |
| DOC / DOCX | Yes | Yes |
| JSON | Yes | Yes |
| XML | Yes | Yes |
| SQL / SQLite | Yes | Yes |
| Chart PNG | No | Yes |

## Limitations

- The application is a local desktop tool and not a production ETL platform.
- PDF and DOC text extraction depends on the source file being text-based and readable.
- SQL support is limited to SQLite-compatible inputs.
- Large or very complex files may require validation against the source data.
- The app stores imported and processed data locally in a SQLite workspace database.

## Notes

FrMat is intended as a practical local data review and cleanup project. It may be used as a base for further enhancements such as expanded file support, more advanced analysis tools, or improved packaging and distribution.

## License

This project is licensed under the MIT License. See `LICENSE` for details.
