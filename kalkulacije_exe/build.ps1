# Build skripta za kalk_excel.exe
# - Vir: kalk_excel.py (v tem folderju)
# - Cilj: c:\Projekti\ai_exe_dll\kalkulacije_exe\kalk_excel.exe
# - Potrebuje: Python 3, pip install --user openpyxl pyinstaller
#
# Zagon iz PowerShell-a:
#   .\build.ps1

$ErrorActionPreference = "Stop"

$here = $PSScriptRoot
$out  = "C:\Projekti\ai_exe_dll\kalkulacije_exe"

Write-Host "== kalk_excel.exe build ==" -ForegroundColor Cyan
Write-Host "Source: $here"
Write-Host "Output: $out"

if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out -Force | Out-Null }

$build = Join-Path $here "build"
$dist  = Join-Path $here "dist"
if (Test-Path $build) { Remove-Item $build -Recurse -Force }
if (Test-Path $dist)  { Remove-Item $dist  -Recurse -Force }

python -m PyInstaller `
    --onefile `
    --console `
    --name kalk_excel `
    --distpath $out `
    --workpath $build `
    --specpath $build `
    --clean `
    --noconfirm `
    (Join-Path $here "kalk_excel.py")

if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed" }

if (Test-Path $build) { Remove-Item $build -Recurse -Force }

$exe = Join-Path $out "kalk_excel.exe"
if (Test-Path $exe) {
    $size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host "OK: $exe ($size MB)" -ForegroundColor Green
} else {
    throw "Build dokoncan ampak .exe ni v $out"
}
