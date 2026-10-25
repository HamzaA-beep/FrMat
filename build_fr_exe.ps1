# Place this script, Fr.png, data_workbench.py, requirements.txt, and
# data_workbench_theme.qss in the same directory. Run on Windows with Python installed.
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$png = Join-Path $PSScriptRoot 'Fr.png'
$ico = Join-Path $PSScriptRoot 'Fr.ico'
if (-not (Test-Path $png)) { throw 'Download the Fr icon image as Fr.png into this folder first.' }

# ICO can embed a 256px PNG directly; resize the generated image if needed.
Add-Type -AssemblyName System.Drawing
$image = [System.Drawing.Image]::FromFile($png)
$bitmap = [System.Drawing.Bitmap]::new(256, 256)
try {
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.DrawImage($image, 0, 0, 256, 256)
    } finally { $graphics.Dispose() }
    $memory = [System.IO.MemoryStream]::new()
    try {
        $bitmap.Save($memory, [System.Drawing.Imaging.ImageFormat]::Png)
        $bytes = $memory.ToArray()
    } finally { $memory.Dispose() }
} finally { $bitmap.Dispose(); $image.Dispose() }
$stream = [System.IO.File]::Create($ico)
$writer = [System.IO.BinaryWriter]::new($stream)
try {
    $writer.Write([uint16]0)       # Reserved
    $writer.Write([uint16]1)       # Icon type
    $writer.Write([uint16]1)       # One image
    $writer.Write([byte]0)         # 0 denotes width 256
    $writer.Write([byte]0)         # 0 denotes height 256
    $writer.Write([byte]0)         # No palette
    $writer.Write([byte]0)         # Reserved
    $writer.Write([uint16]1)       # Color planes
    $writer.Write([uint16]32)      # Bits per pixel
    $writer.Write([uint32]$bytes.Length)
    $writer.Write([uint32]22)      # Header (6) + directory entry (16)
    $writer.Write($bytes)
} finally { $writer.Dispose() }

py -m pip install -r requirements.txt pyinstaller
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
py -m PyInstaller --noconfirm --clean --onefile --windowed --name DataWorkbench --icon $ico --add-data 'data_workbench_theme.qss;.' data_workbench.py
if ($LASTEXITCODE -ne 0) { throw 'EXE build failed.' }
Write-Host "Built $PSScriptRoot\dist\DataWorkbench.exe with the Fr icon."
Write-Host 'To brand an NSIS Setup.exe too, add Icon "Fr.ico" and UninstallIcon "Fr.ico" to its .nsi script.'
