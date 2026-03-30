#!/bin/bash
# 案件登録サーバー 起動スクリプト
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "================================================"
echo " 📁 案件登録ツール サーバー起動"
echo "================================================"
echo ""

# Python3の場所を確認
PYTHON=""
for p in python3 /usr/bin/python3 /usr/local/bin/python3 /opt/homebrew/bin/python3; do
  if command -v "$p" &>/dev/null; then
    PYTHON="$p"
    break
  fi
done

if [ -z "$PYTHON" ]; then
  echo "❌ Python3がインストールされていません"
  read -p "Enterで閉じる..."
  exit 1
fi

if [ ! -f "$SCRIPT_DIR/notion_server.py" ]; then
  echo "❌ notion_server.py が見つかりません: $SCRIPT_DIR"
  read -p "Enterで閉じる..."
  exit 1
fi

echo "✅ $($PYTHON --version)"

# 必要パッケージの確認・自動インストール
echo "📦 必要パッケージを確認中..."
"$PYTHON" -c "import openpyxl" 2>/dev/null || {
  echo "  → openpyxl をインストール中..."
  "$PYTHON" -m pip install openpyxl --break-system-packages --quiet 2>/dev/null \
    || "$PYTHON" -m pip install openpyxl --quiet 2>/dev/null
  echo "  → インストール完了"
}

echo "🚀 サーバーを起動します..."
echo ""

# 既存のサーバープロセスを終了（ポート8765を使っているものを強制終了）
OLD_PID=$(lsof -ti tcp:8765 2>/dev/null)
if [ -n "$OLD_PID" ]; then
  echo "⚠️  既存サーバー(PID:$OLD_PID)を終了します..."
  kill "$OLD_PID" 2>/dev/null
  sleep 1
fi

# サーバーをバックグラウンドで起動
"$PYTHON" "$SCRIPT_DIR/notion_server.py" &
SERVER_PID=$!

# 起動待ち（最大5秒）
for i in 1 2 3 4 5; do
  sleep 1
  if curl -s http://127.0.0.1:8765/api/health > /dev/null 2>&1; then
    break
  fi
done

# ブラウザで自動オープン
echo "🌐 ブラウザを開きます..."
open "http://localhost:8765/"

echo ""
echo "  ▶ 案件登録: http://localhost:8765/案件登録ツール.html"
echo "  ▶ 顧客登録: http://localhost:8765/顧客登録ツール.html"
echo ""
echo "  停止するには Ctrl+C を押してください"
echo ""

# フォアグラウンドで待機
wait $SERVER_PID
