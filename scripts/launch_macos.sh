#!/bin/sh
# このコピーの仮想環境だけを使用する。環境構築・外部通信は行わない。
set -eu

PROJECT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
cd "$PROJECT_DIR"

if [ ! -x "$PROJECT_DIR/.venv/bin/python" ]; then
    echo "このフォルダ用のPython環境（.venv）が見つかりません。" >&2
    echo "ターミナルでこのフォルダを開き、READMEの環境構築手順を実行してください。" >&2
    echo "uv sync --python 3.12 --frozen --extra dev" >&2
    echo "通常の起動では、環境の作成・変更やダウンロードを行いません。" >&2
    exit 1
fi

exec "$PROJECT_DIR/.venv/bin/python" -m sinta_shift
