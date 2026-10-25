# FrMat packaging and installation

FrMat is a local desktop application for working with data files, cleaning them, and exporting the final result without relying on a remote service.

This repository includes the source code, dependency list, packaging notes, and Windows install scripts required to run the application locally on a machine.

## Requirements

- Python 3.11 or 3.12 (64-bit recommended)
- Windows is the primary supported platform
- Internet access is needed to install the required Python packages
- Optional: NSIS for building a standalone Windows installer

## Install from source

1. Install Python and ensure `py` is available in PATH.
2. Open PowerShell in the project folder.
3. Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\Install-From-Source.ps1
```

This creates a local virtual environment and installs the dependencies required by FrMat.

## Run the application

After installation, launch the app with:

```powershell
python data_workbench.py
```

Or directly from the environment used by the installer.

## Build a standalone Windows installer

On a Windows machine with Python and NSIS installed, run the build script:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```

The build output is written to the `dist` directory.

## Package and release notes

- The repo is intended for local Windows use and packaging workflows.
- `.doc` import may require LibreOffice to be installed and available on `PATH` as `soffice`.
- The application is meant for local-first, privacy-conscious analysis rather than cloud-based collaboration.

## Repository layout

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
