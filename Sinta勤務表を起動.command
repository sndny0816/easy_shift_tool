#!/bin/sh
PROJECT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
sh "$PROJECT_DIR/scripts/launch_macos.sh"
RESULT=$?
if [ "$RESULT" -ne 0 ]; then
    echo "起動できませんでした。上に表示された理由とREADMEを確認してください。" >&2
    if [ -t 0 ]; then
        printf 'Enterキーで閉じます。'
        read -r ANSWER
    fi
fi
exit "$RESULT"
