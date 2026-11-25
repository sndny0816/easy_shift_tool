$ErrorActionPreference = "Stop"

if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
    throw "Windows上で実行してください。macOS版はmacOS上でビルドします。"
}

$ProjectDir = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectDir
$PythonPath = Join-Path $ProjectDir ".venv\Scripts\python.exe"
if (-not (Test-Path $PythonPath)) {
    throw "先にREADMEの手順で.venvと開発用依存関係を用意してください。"
}

& $PythonPath -c "import tkinter; import PyInstaller; import pandas; import openpyxl; import ortools"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $PythonPath -m PyInstaller --noconfirm sinta_shift.spec
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Host "ビルド完了: dist\SintaShift\SintaShift.exe"
Write-Host "dist\SintaShiftフォルダ全体を配布し、対象PCでサンプル計算とExcel保存を確認してください。"
