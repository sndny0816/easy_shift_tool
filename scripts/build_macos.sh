#!/bin/sh
set -eu

PROJECT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
cd "$PROJECT_DIR"

if [ "$(uname -s)" != "Darwin" ]; then
    echo "macOS上で実行してください。Windows版はWindows上でビルドします。" >&2
    exit 1
fi
if [ ! -x .venv/bin/python ]; then
    echo "先にREADMEの手順で.venvと開発用依存関係を用意してください。" >&2
    exit 1
fi

.venv/bin/python -c 'import tkinter; import PyInstaller; import pandas; import openpyxl; import ortools'
.venv/bin/python -m PyInstaller --noconfirm sinta_shift.spec
echo "ビルド完了: dist/Sinta Shift.app"
echo "対象Macで実際に起動し、サンプルの計算とExcel保存を確認してください。"
