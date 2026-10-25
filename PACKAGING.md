# Windows installation and GitHub packaging

This folder is a GitHub-ready source repository. It contains the original application, dependency list and README, plus two installation routes. No precompiled Windows `.exe` is included in the ZIP: a Windows binary cannot be produced or tested in this environment.

## Install now from source (Windows)

1. Install **Python 3.12, 64-bit**, including the Python launcher (`py`); internet access is required to download packages.
2. Extract this ZIP, open PowerShell in the extracted `DataWorkbench_GitHub` folder, and run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\installer\Install-From-Source.ps1
   ```
3. Launch **Data Workbench (Source)** from the Start menu. The script installs the project under `%LOCALAPPDATA%\Programs\DataWorkbench-Source`, without administrator privileges. Keep that directory if you want to use the app. The workspace lives separately in `%LOCALAPPDATA%\DataWorkbench`.

The script installs Python dependencies rather than creating a self-contained executable. Windows may require additional VC++ runtime components, depending on installed packages.

## Build a standalone Windows installer

On a Windows PC with Python 3.12 (64-bit), internet access, and [NSIS](https://nsis.sourceforge.io/) installed, run from the repository root:
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```
The output is `dist\DataWorkbenchSetup.exe`. It installs per user, creates Start menu shortcuts, registers an uninstaller, and leaves your workspace untouched on uninstall. The executable is unsigned, so Windows SmartScreen may warn; review the source and build it yourself before running it.

## Build through GitHub

Create a repository, upload the contents of this folder (not the outer ZIP), then use **Actions → Build Windows installer → Run workflow**. Download the `DataWorkbench-Windows-Installer` artifact from the completed run; extract it to get `DataWorkbenchSetup.exe`. A push of a `v*` tag also starts the build. GitHub Actions must be enabled, and the run requires access to PyPI and Chocolatey. This workflow uploads an artifact; it does not automatically publish a GitHub Release.

The compiled installer and its first-run behavior have **not been tested on Windows here**. `.doc` import still requires a separate LibreOffice installation available as `soffice` on PATH. See README.md for supported formats and data-processing limitations.
