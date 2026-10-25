# FrMat installation guide

This guide explains how to install and run FrMat on a Windows machine.

## Requirements

- Python 3.11 or 3.12
- 64-bit Python recommended
- Access to `pip` to install dependencies
- A local workstation or laptop with enough disk space

## Install from source

Open PowerShell in the project directory and run:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python data_workbench.py
```

If Python 3.11 is installed, replace `-3.12` with `-3.11`.

## Optional application packaging

To build a standalone Windows executable with PyInstaller:

```powershell
py -3.12 -m pip install --upgrade pip pyinstaller
pyinstaller --noconfirm --onefile --windowed --name FrMat --collect-all matplotlib --collect-all pandas data_workbench.py
```

The output appears in the `dist` folder.

## Notes

- `.doc` imports may require LibreOffice and a callable `soffice` executable on PATH.
- For best results, keep the project folder and generated environment on the same machine.
- The app stores its local workspace under the user profile’s local app data directory.

## Summary

FrMat is designed to be easy to install and run locally, without requiring a server or cloud account.
