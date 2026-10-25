# FrMat installation and packaging

This repository contains the source files for FrMat, including the application script, dependency list, and packaging guidance. The project is designed to run locally on Windows without requiring a cloud service.

## Install from source

1. Install Python 3.11 or 3.12 (64-bit) and ensure `py` is available in PATH.
2. Open PowerShell in the project folder.
3. Run the installer script:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\Install-From-Source.ps1
```

This creates a local virtual environment and installs the dependencies required by the application.

## Build a standalone Windows installer

On a Windows PC with Python and NSIS installed, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```

The resulting installer is created in the `dist` folder.

## GitHub workflow usage

The repository is also structured for GitHub-hosted builds or release packaging. Use the existing workflow or build scripts if you want to package the project for distribution.

## Notes

- `.doc` import may require LibreOffice to be installed and available on `PATH` as `soffice`.
- This repository is intended for local Windows use and packaging workflows rather than server-side deployment.
