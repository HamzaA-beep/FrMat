# Place this script, Fr.png, data_workbench.py, requirements.txt, and
# data_workbench_theme.qss in the same directory.
# Run with PowerShell on Windows.
#
# Example:
#   .\build_fr_exe.ps1

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

$png = Join-Path $PSScriptRoot 'Fr.png'
$ico = Join-Path $PSScriptRoot 'Fr.ico'

if (-not (Test-Path $png)) {
    throw "Missing Fr.png. Put your image beside this script as Fr.png."
}

# Install Pillow if needed for reliable .ico generation.
& $pythonCmd -m pip install pillow --quiet
if ($LASTEXITCODE -ne 0) {
    throw "Failed to install Pillow."
}

# Create a proper ICO file from PNG.
& $pythonCmd -c "from PIL import Image; img = Image.open(r'$png'); img = img.convert('RGBA'); img.save(r'$ico', format='ICO', sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)])"
if ($LASTEXITCODE -ne 0) {
    throw "Failed to create Fr.ico from Fr.png."
}

# Install Python dependencies and PyInstaller
& $pythonCmd -m pip install -r requirements.txt pyinstaller --quiet
if ($LASTEXITCODE -ne 0) {
    throw "Dependency installation failed."
}

# Build the executable with app icon
& $pythonCmd -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name FrMat `
    --icon "$ico" `
    --add-data "data_workbench_theme.qss;." `
    "data_workbench.py"

if ($LASTEXITCODE -ne 0) {
    throw "EXE build failed."
}

Write-Host "Built: $PSScriptRoot\dist\FrMat.exe"
Write-Host "Icon: $ico"
Write-Host "If you want a setup installer later, use this EXE as the main app."
