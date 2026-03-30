#!/bin/bash
# =====================================================
#  業務ポータルサーバー ── 自動起動インストーラー
#  このファイルをダブルクリックするだけでOK。
#  Mac起動時にサーバーが自動で立ち上がるようになります。
# =====================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PLIST_LABEL="com.notsuo.ankenserver"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"
SERVER_PY="$SCRIPT_DIR/notion_server.py"
LOG_DIR="$HOME/Library/Logs/AnkenServer"
USER_ID=$(id -u)

echo "================================================"
echo " 🚀 業務ポータル 自動起動インストーラー"
echo "================================================"
echo ""

# ── Python3 を探す ──
PYTHON=""
for p in /opt/homebrew/bin/python3 /usr/local/bin/python3 python3 /usr/bin/python3; do
  if command -v "$p" &>/dev/null; then
    PYTHON="$(command -v "$p")"
    break
  fi
done

if [ -z "$PYTHON" ]; then
  echo "❌ Python3 が見つかりません。"
  echo "   https://www.python.org からインストールしてください。"
  read -p "Enterで閉じる..."
  exit 1
fi
echo "✅ Python: $PYTHON ($($PYTHON --version 2>&1))"

# ── openpyxl インストール ──
echo "📦 openpyxl を確認中..."
"$PYTHON" -c "import openpyxl" 2>/dev/null || {
  echo "  → インストール中..."
  "$PYTHON" -m pip install openpyxl --break-system-packages --quiet 2>/dev/null \
    || "$PYTHON" -m pip install openpyxl --quiet 2>/dev/null
  echo "  ✅ openpyxl インストール完了"
}

# ── ログディレクトリ作成 ──
mkdir -p "$LOG_DIR"

# ── 既存の LaunchAgent を停止・削除 ──
echo "♻️  既存の設定を確認中..."
# macOS 13+ の新方式
launchctl bootout "gui/${USER_ID}/${PLIST_LABEL}" 2>/dev/null
# 旧方式（念のため）
launchctl unload "$PLIST_PATH" 2>/dev/null
sleep 1

# ── plist 作成 ──
cat > "$PLIST_PATH" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${PLIST_LABEL}</string>

  <key>ProgramArguments</key>
  <array>
    <string>${PYTHON}</string>
    <string>${SERVER_PY}</string>
  </array>

  <!-- Mac起動・ログイン時に自動スタート -->
  <key>RunAtLoad</key>
  <true/>

  <!-- クラッシュ・終了しても自動再起動 -->
  <key>KeepAlive</key>
  <true/>

  <!-- 作業ディレクトリ -->
  <key>WorkingDirectory</key>
  <string>${SCRIPT_DIR}</string>

  <!-- ログ出力先 -->
  <key>StandardOutPath</key>
  <string>${LOG_DIR}/server.log</string>
  <key>StandardErrorPath</key>
  <string>${LOG_DIR}/server_err.log</string>
</dict>
</plist>
PLIST

echo "📝 設定ファイルを作成: $PLIST_PATH"

# ── ロード（macOS 13+ 対応）──
# 新方式で試して、ダメなら旧方式にフォールバック
if ! launchctl bootstrap "gui/${USER_ID}" "$PLIST_PATH" 2>/dev/null; then
  launchctl load "$PLIST_PATH" 2>/dev/null
fi

echo "⏳ サーバー起動を待っています..."
sleep 3

# ── 起動確認 ──
if curl -s --max-time 5 http://127.0.0.1:8765/api/health > /dev/null 2>&1; then
  # ローカルIPを取得
  LOCAL_IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "（IPアドレスを確認してください）")

  echo ""
  echo "================================================"
  echo " ✅ インストール完了！"
  echo "================================================"
  echo ""
  echo "  📱 iPhone からアクセス（同じWiFiまたはTailscale）:"
  echo "     http://${LOCAL_IP}:8765"
  echo ""
  echo "  💻 Mac からアクセス:"
  echo "     http://localhost:8765"
  echo ""
  echo "  次回からMacを起動するだけで自動的に立ち上がります。"
  echo "  （手動で起動.command を実行する必要はありません）"
  echo ""
  # ブラウザで開く
  open "http://localhost:8765/"
else
  echo ""
  echo "⚠️  起動を確認できませんでした。ログを確認してください:"
  echo "   $LOG_DIR/server_err.log"
  echo ""
  echo "   ログを開きますか？"
  if [ -f "$LOG_DIR/server_err.log" ]; then
    cat "$LOG_DIR/server_err.log" | tail -20
  fi
fi

echo ""
echo "── アンインストールしたい場合 ──"
echo "  launchctl bootout gui/${USER_ID}/${PLIST_LABEL}"
echo "  rm \"$PLIST_PATH\""
echo ""
read -p "Enterで閉じる..."
