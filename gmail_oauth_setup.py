#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gmail API 用のリフレッシュトークンを1回だけ作る（自分のMacで実行）。

使い方:
  python3 gmail_oauth_setup.py
  → クライアントID と クライアントシークレット を聞かれるので貼る
  → ブラウザが開くので Gmail アカウントでログインして「許可」
  → 画面に出た3行を Railway の Variables に入れる
"""
import http.server, json, threading, urllib.parse, urllib.request, webbrowser, secrets, sys

SCOPE = "https://www.googleapis.com/auth/gmail.send"
PORT = 8921

def main():
    cid = input("クライアントID: ").strip()
    csec = input("クライアントシークレット: ").strip()
    if not cid or not csec:
        print("両方必要です"); sys.exit(1)
    redirect = f"http://localhost:{PORT}/"
    state = secrets.token_urlsafe(16)
    url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode({
        "client_id": cid, "redirect_uri": redirect, "response_type": "code", "scope": SCOPE,
        "access_type": "offline", "prompt": "consent", "state": state})
    got = {}
    class H(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            if q.get("state", [""])[0] != state or "code" not in q:
                self.send_response(400); self.end_headers(); self.wfile.write("NG".encode()); return
            got["code"] = q["code"][0]
            self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8"); self.end_headers()
            self.wfile.write("<h2>許可を受け取りました。この画面は閉じて、ターミナルに戻ってください。</h2>".encode())
        def log_message(self, *a): pass
    srv = http.server.HTTPServer(("localhost", PORT), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    print("\nブラウザが開きます。Gmail のアカウントでログインして「許可」を押してください。")
    print("（「このアプリは確認されていません」と出たら「詳細」→「（アプリ名）に移動」で進めます）\n")
    webbrowser.open(url)
    while "code" not in got:
        threading.Event().wait(0.5)
    srv.shutdown()
    body = urllib.parse.urlencode({"code": got["code"], "client_id": cid, "client_secret": csec,
                                   "redirect_uri": redirect, "grant_type": "authorization_code"}).encode()
    req = urllib.request.Request("https://oauth2.googleapis.com/token", data=body,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    d = json.loads(urllib.request.urlopen(req, timeout=30).read())
    if "refresh_token" not in d:
        print("リフレッシュトークンが返ってきませんでした:", d); sys.exit(1)
    print("\n===== Railway の Variables にこの3つを入れてください =====")
    print(f"GMAIL_OAUTH_CLIENT_ID={cid}")
    print(f"GMAIL_OAUTH_CLIENT_SECRET={csec}")
    print(f"GMAIL_OAUTH_REFRESH_TOKEN={d['refresh_token']}")
    print("==========================================================")
    print("※ Google Cloud の OAuth 同意画面が「テスト」のままだと、このトークンは7日で切れます。「本番環境に公開」にしておいてください。")

if __name__ == "__main__":
    main()
