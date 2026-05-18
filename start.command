#!/bin/bash
# USX 運転記録表 起動スクリプト（Mac用）
# このファイルをダブルクリックして起動してください

cd "$(dirname "$0")"

echo "================================================"
echo " USX 単体運転記録表 起動中..."
echo " ブラウザが自動で開きます"
echo " 終了するにはこのウィンドウを閉じてください"
echo "================================================"

# Pythonの確認
if ! command -v python3 &> /dev/null; then
    osascript -e 'display alert "Pythonが見つかりません" message "python.org からPython 3をインストールしてください。"'
    exit 1
fi

# ライブラリインストール（初回のみ）
echo "必要なライブラリを確認中..."
python3 -m pip install flask openpyxl --quiet --exists-action i

# サーバー起動
python3 server.py
