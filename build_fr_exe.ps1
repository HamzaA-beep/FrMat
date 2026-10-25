# Place this script, data_workbench.py, requirements.txt, and
# data_workbench_theme.qss in the same directory.
# Run with PowerShell on Windows.

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$pythonCmd = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonCmd = "py"
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCmd = "python"
} else {
    throw "Python is not installed or not on PATH."
}

& $pythonCmd -m pip install -r requirements.txt pyinstaller --quiet
if ($LASTEXITCODE -ne 0) {
    throw "Dependency installation failed."
}

& $pythonCmd -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name FrMat `
    --add-data "data_workbench_theme.qss;." `
    "data_workbench.py"

if ($LASTEXITCODE -ne 0) {
    throw "EXE build failed."
}

Write-Host "Built: $PSScriptRoot\dist\FrMat.exe"
