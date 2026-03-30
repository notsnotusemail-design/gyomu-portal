"""
案件登録サーバー
HTMLフォームからNotionへ直接登録するためのローカルAPIサーバー

起動方法:
  python server.py

初回のみ: Notionトークンを下記に設定してください
"""

import json
import os
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse
import urllib.request
import urllib.error

# ============================================================
# ★ ここにNotionのIntegration Tokenを貼り付けてください ★
# 取得方法: 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
# ============================================================
NOTION_TOKEN = "YOUR_NOTION_TOKEN_HERE"

# データベースID（変更不要）
CASE_DATABASE_ID = "2e13288b84b081faa8e8fb22589ce378"  # 案件表

# お客様ページID（重要備考の更新用）
CUSTOMER_PAGE_IDS = {
    "201": "3293288b84b080dc8adff83cb6b3b9a2",
    "301": "2593288b84b0803b9dddf188bb63d703",
    "302": "2743288b84b08096b828ed6983716d84",
    "304": "2d53288b84b08061bc53c01405d0a4e8",
    "305": "2743288b84b08067bc16ccba2966f49a",
    "307": "2ed3288b84b080179aeffa72b1ff0ba4",
    "308": "2f43288b84b08005b2d1ec83b3d8d5b3",
    "309": "2f63288b84b08046800bccb41768a181",
    "311": "31f3288b84b080baaaaec56eec3faa68",
    "315": "31a3288b84b0805a9dfde2285627b489",
    "316": "e937740806cd406fa11fcf972605078d",
    "3202-302": "31f3288b84b080db8ed3d56363e102ef",
    "3302-302": "31a3288b84b080c9a868d2198a02784c",
}

NOTION_VERSION = "2022-06-28"
PORT = 5055

def notion_request(method, path, body=None):
    url = f"https://api.notion.com/v1/{path}"
    headers = {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }
    data = json.dumps(body).encode("utf-8") if body else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as res:
            return json.loads(res.read().decode("utf-8")), None
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8")
        return None, f"HTTP {e.code}: {err}"
    except Exception as e:
        return None, str(e)

def register_case(data):
    """案件表に新しいエントリを作成する"""
    properties = {
        "当方案件番号": {
            "title": [{"text": {"content": data.get("number", "")}}]
        },
        "お客様no/名": {
            "rich_text": [{"text": {"content": data.get("customerNo", "")}}]
        },
        "進捗": {
            "status": {"name": data.get("progress", "未着手")}
        },
    }

    # 締切日
    if data.get("deadline"):
        properties["案件締切日・進行"] = {"date": {"start": data["deadline"]}}

    # 素材名
    if data.get("materialName"):
        properties["備考/素材名"] = {
            "rich_text": [{"text": {"content": data["materialName"]}}]
        }

    # 単価
    if data.get("price"):
        try:
            properties["単価（単価-外注費）"] = {"number": float(data["price"])}
        except:
            pass

    # 指定ファイル名
    if data.get("fileName"):
        properties["指定案件ファイル名"] = {
            "rich_text": [{"text": {"content": data["fileName"]}}]
        }

    # 備考
    if data.get("memo"):
        properties["備考"] = {
            "rich_text": [{"text": {"content": data["memo"]}}]
        }

    body = {
        "parent": {"database_id": CASE_DATABASE_ID},
        "properties": properties,
    }

    result, err = notion_request("POST", "pages", body)
    return result, err

def update_customer_memo(customer_no, new_case_label, new_case_name):
    """お客様の重要備考に新案件を追記する"""
    page_id = CUSTOMER_PAGE_IDS.get(str(customer_no))
    if not page_id:
        return None, f"お客様No.{customer_no}のページIDが見つかりません"

    # 現在の重要備考を取得
    current, err = notion_request("GET", f"pages/{page_id}")
    if err:
        return None, err

    current_text = ""
    try:
        props = current.get("properties", {})
        rt = props.get("重要備考", {}).get("rich_text", [])
        current_text = "".join([t.get("text", {}).get("content", "") for t in rt])
    except:
        pass

    # 追記
    append_text = f"\n{new_case_label}案件：{new_case_name}"
    new_text = current_text.rstrip() + append_text

    body = {
        "properties": {
            "重要備考": {
                "rich_text": [{"text": {"content": new_text}}]
            }
        }
    }

    result, err = notion_request("PATCH", f"pages/{page_id}", body)
    return result, err


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        # ログを簡潔に
        print(f"[{self.command}] {self.path} - {args[0] if args else ''}")

    def send_cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors()
        self.end_headers()

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length).decode("utf-8")) if length else {}

        response = {"ok": False, "message": ""}

        if parsed.path == "/register":
            # 案件登録
            if NOTION_TOKEN == "YOUR_NOTION_TOKEN_HERE":
                response = {"ok": False, "message": "Notionトークンが設定されていません。server.py の NOTION_TOKEN を設定してください。"}
            else:
                result, err = register_case(body)
                if err:
                    response = {"ok": False, "message": f"登録失敗: {err}"}
                else:
                    page_url = result.get("url", "")
                    response = {"ok": True, "message": "登録完了！", "url": page_url}

                    # 新案件追記が必要な場合
                    if body.get("isNewCase") and body.get("newCaseLabel") and body.get("newCaseName"):
                        _, err2 = update_customer_memo(
                            body.get("customerNo"),
                            body.get("newCaseLabel"),
                            body.get("newCaseName")
                        )
                        if err2:
                            response["message"] += f"（顧客情報の更新に失敗: {err2}）"
                        else:
                            response["message"] += f" + 顧客情報に{body['newCaseLabel']}案件を追記しました。"

        elif parsed.path == "/health":
            response = {"ok": True, "message": "サーバー起動中"}

        else:
            response = {"ok": False, "message": "不明なエンドポイント"}

        resp_bytes = json.dumps(response, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_cors()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(resp_bytes)))
        self.end_headers()
        self.wfile.write(resp_bytes)


def main():
    if NOTION_TOKEN == "YOUR_NOTION_TOKEN_HERE":
        print("=" * 60)
        print("⚠️  Notionトークンが未設定です")
        print()
        print("設定方法:")
        print("1. https://www.notion.so/my-integrations を開く")
        print("2. 「新しいインテグレーション」を作成")
        print("3. 生成されたトークン（secret_...）をコピー")
        print("4. このファイル（server.py）の NOTION_TOKEN に貼り付け")
        print("5. 案件進行管理のページを開いてIntegrationと接続する")
        print("=" * 60)

    server = HTTPServer(("localhost", PORT), Handler)
    print(f"\n✅ 案件登録サーバー起動中 → http://localhost:{PORT}")
    print("停止するには Ctrl+C を押してください\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nサーバーを停止しました")
        server.server_close()
        sys.exit(0)

if __name__ == "__main__":
    main()
