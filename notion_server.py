#!/usr/bin/env python3
# coding: utf-8
"""
案件登録サーバー
NotionのAPIへ直接登録します。
起動: python3 notion_server.py
"""

import sys, json, os, re, tempfile, datetime, uuid, hashlib, io
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.request import urlopen, Request
from urllib.error import HTTPError, URLError
from urllib.parse import unquote
import json as jsonlib
import email, email.parser
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False
    print("⚠️  openpyxlが未インストール。請求書xlsx生成を使うには: pip install openpyxl --break-system-packages")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# .envファイルから環境変数を読み込む（Macローカル用）
_env_file = os.path.join(SCRIPT_DIR, ".env")
if os.path.exists(_env_file):
    with open(_env_file) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))

# ============================================================
# Notionトークンは環境変数 NOTION_TOKEN から読み込みます
# Railway: Variables タブで設定
# Mac ローカル: notion_server.py と同じフォルダに .env ファイルを作成し
#              NOTION_TOKEN=ntn_xxxxx と記載
# ============================================================
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")

# ポート番号（Railway等クラウドは$PORT環境変数を使用）
PORT = int(os.environ.get("PORT", 8765))

# NotionのデータベースID（案件表）
CASE_DB_ID = "2e13288b-84b0-81fa-a8e8-fb22589ce378"

# お客様データベースID
CUSTOMER_DB_ID = "1513288b-84b0-8035-ae58-d410686d282d"

# 予定データベースID
SCHEDULE_DB_ID = "4a2a3ad54b5b41a6b138420ee5841ee3"

# 引き継ぎデータベースID（Notionで管理）
HANDOVER_DB_ID = "92c91778-a575-445b-80b8-f233a0c23261"

# メール定型文データベースID（Notionで管理）
MAIL_TEMPLATE_DB_ID = os.environ.get("MAIL_TEMPLATE_DB_ID", "3df3288b-84b0-814e-a04b-ddcb121341e9")

# 日報（カフェで作業した日の記録・経費メモ用）データベースID（Notionで管理）
NIPPO_DB_ID = os.environ.get("NIPPO_DB_ID", "3fc3288b-84b0-81ed-9387-c8f3ac7f1adf")

# 営業ツール（フォーム営業・メール営業）
EIGYO_LIST_DB_ID = os.environ.get("EIGYO_LIST_DB_ID", "3fd3288b-84b0-8198-a5ad-faa5a26dcb72")
EIGYO_TPL_DB_ID  = os.environ.get("EIGYO_TPL_DB_ID",  "3fd3288b-84b0-811e-b1e9-ea8cc7700108")
# メール営業の送信元（Gmail）。アプリパスワードは Railway の環境変数にだけ置く
GMAIL_USER         = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "")
EIGYO_SENDER_NAME  = os.environ.get("EIGYO_SENDER_NAME", "野津 欧")
EIGYO_DAILY_LIMIT  = int(os.environ.get("EIGYO_DAILY_LIMIT", "40"))   # 1日に送るメールの上限
EIGYO_STATUSES = ["未着手", "送信済", "返信あり", "商談", "保留", "送信失敗", "営業お断り", "対象外"]
# 送り主の情報（メール本文の末尾に付く。特定電子メール法：氏名・住所・配信停止の案内が必要）
EIGYO_SENDER_TEL     = os.environ.get("EIGYO_SENDER_TEL", "")
EIGYO_SENDER_URL     = os.environ.get("EIGYO_SENDER_URL", "")
EIGYO_SENDER_ADDRESS = os.environ.get("EIGYO_SENDER_ADDRESS", "")
# 自動送信（サーバーが営業リストの「未着手」を順に送る）
EIGYO_AUTO_DEFAULT = os.environ.get("EIGYO_AUTO", "0") == "1"          # 起動時にONにするか
EIGYO_AUTO_HOURS   = os.environ.get("EIGYO_AUTO_HOURS", "10-18")       # 送る時間帯（JST・開始-終了）
EIGYO_AUTO_DAYS    = os.environ.get("EIGYO_AUTO_DAYS", "1-5")          # 送る曜日（1=月 … 7=日）
EIGYO_AUTO_PER_DAY = int(os.environ.get("EIGYO_AUTO_PER_DAY", str(EIGYO_DAILY_LIMIT)))  # 自動で送る1日の上限
EIGYO_AUTO_GAP_SEC = int(os.environ.get("EIGYO_AUTO_GAP_SEC", "480"))  # 送信と送信の間隔（秒）
EIGYO_AUTO_DRY     = os.environ.get("EIGYO_AUTO_DRY", "0") == "1"      # 1なら送ったことにして実際は送らない（テスト用）
EIGYO_LOW_WATER    = int(os.environ.get("EIGYO_LOW_WATER", "20"))      # 残りがこれを切ったら「リスト補充」を引き継ぎに書く
EIGYO_AUTO_STATE_FILE = os.path.join(SCRIPT_DIR, "eigyo_auto.json")

# 日次スケジュール保存ファイル（ローカルJSON）
DAILY_SCHEDULE_FILE = os.path.join(SCRIPT_DIR, "daily_schedules.json")

# ── 請求書システム（Notion DB）──────────────────────────
INVOICE_DB_ID = os.environ.get("INVOICE_DB_ID", "25f803a762fc4a4b8c876c8756b52b66")
WORKER_DB_ID  = os.environ.get("WORKER_DB_ID",  "17706159d9854ad4832a6e80a14e285f")

# ── 管理者パスワード認証 ────────────────────────────────
# パスワード変更方法：.env ファイルに ADMIN_PASSWORD=新しいパスワード と記載
# Railway Variables に ADMIN_PASSWORD=新しいパスワード でも可
# デフォルトパスワード: nts2026
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")
# 認証トークン = sha256(パスワード + salt)
def _make_admin_token(pw):
    return hashlib.sha256((pw + "nts_admin_2026").encode()).hexdigest()
# デフォルトトークン = sha256("nts2026" + salt)。ADMIN_PASSWORDが設定されていれば上書き
_DEFAULT_ADMIN_TOKEN = "ec31d1b13ecbc5c119f79ae5ea02ce8a3f3265a354acd6063894c2b8bf863c8f"
ADMIN_TOKEN = _make_admin_token(ADMIN_PASSWORD) if ADMIN_PASSWORD else _DEFAULT_ADMIN_TOKEN

# PDF は一時保存（Railwayでは再デプロイで消えるが、請求書データはNotionに永続保存）
INVOICE_PDF_DIR = os.path.join(SCRIPT_DIR, "invoice_pdfs")
os.makedirs(INVOICE_PDF_DIR, exist_ok=True)

def parse_multipart(raw_body, content_type):
    """multipart/form-data をパースして fields, files を返す"""
    # Content-Typeからboundaryを抽出
    ct_str = str(content_type)
    boundary = None
    for part in ct_str.split(";"):
        part = part.strip()
        if part.startswith("boundary="):
            boundary = part[len("boundary="):].strip('"')
            break
    if not boundary:
        return {}, {}

    fields = {}
    files  = {}
    delimiter = ("--" + boundary).encode()
    end_delim  = ("--" + boundary + "--").encode()

    # bodyをパーツに分割
    for chunk in raw_body.split(delimiter):
        if not chunk or chunk.strip() in (b"", end_delim.lstrip(b"--" + boundary.encode())):
            continue
        if chunk.startswith(b"--"):   # 終端
            continue
        # ヘッダーとボディを分離（\r\n\r\n で区切る）
        if b"\r\n\r\n" not in chunk:
            continue
        header_part, body_part = chunk.split(b"\r\n\r\n", 1)
        # 末尾の \r\n を除去
        if body_part.endswith(b"\r\n"):
            body_part = body_part[:-2]

        headers = {}
        for line in header_part.split(b"\r\n"):
            if b":" in line:
                k, v = line.split(b":", 1)
                headers[k.strip().lower().decode()] = v.strip().decode("utf-8", errors="replace")

        disp = headers.get("content-disposition", "")
        if not disp:
            continue

        name, filename = "", ""
        for seg in disp.split(";"):
            seg = seg.strip()
            if seg.startswith("name="):
                name = seg[5:].strip('"')
            elif seg.startswith("filename="):
                filename = seg[9:].strip('"')

        if not name:
            continue
        if filename:
            files[name] = (filename, body_part)
        else:
            fields[name] = body_part.decode("utf-8", errors="replace")

    return fields, files

def load_daily_schedules():
    if os.path.exists(DAILY_SCHEDULE_FILE):
        try:
            with open(DAILY_SCHEDULE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_daily_schedules(data):
    with open(DAILY_SCHEDULE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# 顧客ページID
CUSTOMER_PAGES = {
    "201":       "3293288b-84b0-80dc-8adf-f83cb6b3b9a2",
    "301":       "2593288b-84b0-803b-9ddd-f188bb63d703",
    "302":       "2743288b-84b0-8096-b828-ed6983716d84",
    "307":       "2ed3288b-84b0-8017-9aef-fa72b1ff0ba4",
    "308":       "2f43288b-84b0-8005-b2d1-ec83b3d8d5b3",
    "311":       "31f3288b-84b0-80ba-aaae-c56eec3faa68",
    "315":       "31a3288b-84b0-805a-9dfd-e2285627b489",
    "316":       "e9377408-06cd-406f-a11f-cf972605078d",
    "3202-302":  "31f3288b-84b0-80db-8ed3-d56363e102ef",
    "3302-302":  "31a3288b-84b0-80c9-a868-d2198a02784c",
}

import time as _time
CUSTOMER_NAME_CACHE = {}   # {customer_no: customer_name}  キャッシュ
CUSTOMER_LINK_CACHE = {}   # {customer_no: {name, url, id}}  Notion顧客ページへのリンク（名前キャッシュと同時に構築）
_CUSTOMER_CACHE_AT  = 0.0  # 最終構築時刻（epoch秒）
CUSTOMER_CACHE_TTL  = 300  # 秒。期限切れで再構築し、後から追加した顧客も取り込む

def get_customer_name_map(force=False):
    """全顧客のNo→名前マップを返す（TTL付きキャッシュ。期限切れor強制時に再構築）"""
    global _CUSTOMER_CACHE_AT
    if (not force) and CUSTOMER_NAME_CACHE and (_time.time() - _CUSTOMER_CACHE_AT) < CUSTOMER_CACHE_TTL:
        return CUSTOMER_NAME_CACHE
    new_map = {}
    new_links = {}
    body = {"page_size": 100}
    cursor = None
    while True:
        if cursor: body["start_cursor"] = cursor
        result, _ = notion_request("POST", f"/databases/{CUSTOMER_DB_ID}/query", body)
        if not result: break
        for page in result.get("results", []):
            try:
                props = page["properties"]
                no   = (props["お客様No."]["rich_text"]   or [{}])[0].get("plain_text","").strip()
                name = (props["クライアント名"]["rich_text"] or [{}])[0].get("plain_text","").strip()
                if no and name:
                    new_map[no] = name
                if no:
                    # 顧客ページのURL。Notionが返さない場合はIDから組み立てる
                    url = page.get("url") or f"https://www.notion.so/{page['id'].replace('-', '')}"
                    new_links[no] = {"name": name, "url": url, "id": page["id"]}
            except Exception:
                pass
        if not result.get("has_more"): break
        cursor = result.get("next_cursor")
    # まるごと差し替え（改名・削除も反映）。取得0件のときは既存キャッシュを維持
    if new_map:
        CUSTOMER_NAME_CACHE.clear()
        CUSTOMER_NAME_CACHE.update(new_map)
    if new_links:
        CUSTOMER_LINK_CACHE.clear()
        CUSTOMER_LINK_CACHE.update(new_links)
    _CUSTOMER_CACHE_AT = _time.time()
    print(f"  👥 顧客名キャッシュ構築: {len(CUSTOMER_NAME_CACHE)}件（リンク {len(CUSTOMER_LINK_CACHE)}件）")
    return CUSTOMER_NAME_CACHE

def get_customer_link_map(force=False):
    """全顧客のNo→{name,url,id}を返す（名前キャッシュと同じTTLで構築）"""
    get_customer_name_map(force)
    return CUSTOMER_LINK_CACHE

NOTION_API = "https://api.notion.com/v1"
HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28",
}

# ========== 引き継ぎ（Notion DB）==========
def parse_handover_page(page):
    """NotionページをHandoverアイテムに変換"""
    try:
        props = page["properties"]
        text       = (props["内容"]["title"] or [{}])[0].get("plain_text","").strip()
        typ        = (props["種別"]["select"] or {}).get("name","todo")
        status     = (props["ステータス"]["select"] or {}).get("name","active")
        date       = (props["日付"]["date"] or {}).get("start","")
        start_date = (props["開始日"]["date"] or {}).get("start","")
        end_date   = (props["終了日"]["date"] or {}).get("start","")
        created_at = (props["作成日"]["date"] or {}).get("start","")
        item = {"id": page["id"], "type": typ, "text": text,
                "status": status, "created_at": created_at}
        if typ == "medium":
            item["start_date"] = start_date
            item["end_date"]   = end_date
        else:
            item["date"] = date
        return item
    except Exception as e:
        print(f"  ⚠️ parse_handover_page: {e}")
        return None

def get_active_handover(date_str):
    """指定日にアクティブな引き継ぎをNotionから取得"""
    body = {"filter": {"property": "ステータス", "select": {"equals": "active"}},
            "page_size": 100}
    result, _ = notion_request("POST", f"/databases/{HANDOVER_DB_ID}/query", body)
    if not result:
        return []
    items = []
    for page in result.get("results", []):
        item = parse_handover_page(page)
        if not item:
            continue
        t = item["type"]
        if t == "medium":
            if item.get("start_date","") <= date_str <= item.get("end_date",""):
                items.append(item)
        elif t == "todo":
            if item.get("date","") <= date_str:
                items.append(item)
        else:
            if item.get("date","") == date_str:
                items.append(item)
    return items

def get_done_handover():
    """完了済みの引き継ぎをNotionから取得"""
    body = {"filter": {"property": "ステータス", "select": {"equals": "done"}},
            "sorts": [{"timestamp": "last_edited_time", "direction": "descending"}],
            "page_size": 100}
    result, _ = notion_request("POST", f"/databases/{HANDOVER_DB_ID}/query", body)
    if not result:
        return []
    items = []
    for page in result.get("results", []):
        item = parse_handover_page(page)
        if item:
            items.append(item)
    return items

def case_label(case):
    """請求書の「詳細」に出す文字列＝案件名。

    案件名（Notionの指定案件ファイル名）を最優先。
    未設定のときだけ 備考 → 案件番号 の順で代用する。
    案件番号が出るのは案件名も備考も空の場合だけで、
    請求書の明細が空欄になるのを避けるための最終手段。
    """
    for key in ("projectName", "filename", "materialName", "note"):
        v = (case.get(key) or "").strip() if isinstance(case.get(key), str) else case.get(key)
        if v:
            return v
    return case.get("number", "")


# ========== メール定型文（Notion DB）==========
# Notionの rich_text は1要素あたりの文字数に上限があるため分割して格納する
MAIL_TPL_CHUNK = 1900

def mail_rt_chunks(text):
    """長文を Notion rich_text の要素に分割"""
    text = text or ""
    return [{"text": {"content": text[i:i + MAIL_TPL_CHUNK]}}
            for i in range(0, len(text), MAIL_TPL_CHUNK)]

def mail_rt_plain(props, key):
    """rich_textプロパティを1本の文字列に戻す"""
    return "".join(t.get("plain_text", "")
                   for t in (props.get(key, {}).get("rich_text") or []))

def parse_mail_template_page(page):
    """Notionページをメール定型文アイテムに変換"""
    try:
        props    = page["properties"]
        title_rt = props.get("テンプレート名", {}).get("title") or []
        name     = "".join(t.get("plain_text", "") for t in title_rt).strip()
        return {
            "id":        page["id"],
            "name":      name,
            "subject":   mail_rt_plain(props, "件名"),
            "body":      mail_rt_plain(props, "本文"),
            "memo":      mail_rt_plain(props, "メモ"),
            "category":  (props.get("カテゴリ", {}).get("select") or {}).get("name", ""),
            "order":     props.get("並び順", {}).get("number"),
            "enabled":   props.get("有効", {}).get("checkbox", True),
            "updated":   page.get("last_edited_time", ""),
            "notionUrl": page.get("url", ""),
        }
    except Exception:
        return None


# ========== 日報（Notion DB）==========
def nippo_task_label(t):
    """作業1件を一覧用の1行にする（例：302 山内様／10月分：カット）"""
    who = " ".join(x for x in [t.get("customer", ""), t.get("customerName", "")] if x).strip()
    head = "／".join(x for x in [who, t.get("case", "")] if x)
    body = (t.get("content") or "").strip()
    return f"{head}：{body}" if head and body else (head or body)

def parse_nippo_page(page):
    """Notionページを日報アイテムに変換"""
    try:
        props    = page["properties"]
        title_rt = props.get("作業内容", {}).get("title") or []
        title    = "".join(t.get("plain_text", "") for t in title_rt).strip()
        tasks = []
        raw = mail_rt_plain(props, "タスク").strip()
        if raw:
            try:
                tasks = [t for t in jsonlib.loads(raw) if isinstance(t, dict)]
            except Exception:
                tasks = []
        if not tasks and title:          # タスク欄ができる前の日報
            tasks = [{"customer": "", "customerName": "", "case": "", "content": title}]
        return {
            "id":     page["id"],
            "title":  title,
            "tasks":  tasks,
            "detail": mail_rt_plain(props, "詳細"),
            "date":   ((props.get("作業日", {}).get("date") or {}).get("start") or "")[:10],
            "start":  mail_rt_plain(props, "開始").strip(),
            "end":    mail_rt_plain(props, "終了").strip(),
            "place":  mail_rt_plain(props, "場所").strip(),
        }
    except Exception:
        return None


# ========== 営業ツール（Notion DB）==========
def eigyo_now_jst():
    return datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))

def eigyo_norm_url(u):
    u = (u or "").strip()
    if u and not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    return u

EIGYO_SHARED_DOMAINS = {"instagram.com", "facebook.com", "x.com", "twitter.com", "youtube.com", "tiktok.com",
                        "note.com", "ameblo.jp", "line.me", "lit.link", "linktr.ee", "google.com", "goo.gl",
                        "sites.google.com", "peraichi.com", "jimdofree.com", "fc2.com", "hatenablog.com",
    "x.gd", "bit.ly", "t.co", "tinyurl.com", "lin.ee", "forms.gle", "youtu.be", "amzn.to", "is.gd", "stand.fm", "hatenadiary.com", "wixsite.com", "studio.site"}

def eigyo_dup_groups(items):
    """同じ会社とみなせる行をまとめる（キーが1つでも一致したら同じグループ）"""
    parent = {x["id"]: x["id"] for x in items}
    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]; a = parent[a]
        return a
    owner = {}
    for x in items:
        for k in eigyo_key(x):
            if k in owner:
                parent[find(x["id"])] = find(owner[k])
            else:
                owner[k] = x["id"]
    groups = {}
    for x in items:
        groups.setdefault(find(x["id"]), []).append(x["id"])
    return [g for g in groups.values() if len(g) > 1]

EIGYO_STATUS_RANK = {"営業お断り": 9, "対象外": 8, "商談": 7, "返信あり": 6, "送信済": 5, "保留": 4, "送信失敗": 3, "未着手": 1}

def eigyo_merge(rows):
    """重複行を1つにまとめる。空欄は埋め、ステータスは進んでいる方（お断りは最優先）を残す"""
    rows = sorted(rows, key=lambda r: r.get("created", ""))
    base = dict(rows[0])
    for r in rows[1:]:
        for f in ("company", "industry", "site", "form", "email", "person", "line", "method", "sentAt", "template"):
            if not base.get(f) and r.get(f):
                base[f] = r[f]
        if EIGYO_STATUS_RANK.get(r.get("status"), 0) > EIGYO_STATUS_RANK.get(base.get("status"), 0):
            base["status"] = r["status"]
            for f in ("method", "sentAt", "template"):
                if r.get(f):
                    base[f] = r[f]
        memo = (r.get("memo") or "").strip()
        if memo and memo not in (base.get("memo") or ""):
            base["memo"] = ((base.get("memo") or "").strip() + "\n" + memo).strip()
    return base

def eigyo_account_path(url):
    """SNS・ブログ等の共用ドメインはアカウント部分まで見る（youtube.com/channel/UCxx・instagram.com/aaa）"""
    segs = [x for x in re.sub(r"^https?://[^/]+", "", url.split("?")[0].split("#")[0]).split("/") if x]
    if not segs:
        return ""
    if segs[0] in ("channel", "c", "user", "u", "p", "in", "pages", "people", "profile") and len(segs) > 1:
        return segs[0] + "/" + segs[1]
    return segs[0]

def eigyo_key(x):
    """重複判定用のキー（サイトのドメイン・メール・会社名）"""
    keys = []
    site = eigyo_norm_url(x.get("site")).lower()
    m = re.search(r"https?://(?:www\.)?([^/?#]+)", site)
    if m:
        dom = m.group(1)
        if dom in EIGYO_SHARED_DOMAINS:
            seg = eigyo_account_path(site)
            if seg:
                keys.append("d:" + dom + "/" + seg)
        else:
            keys.append("d:" + dom)
    if x.get("email"):
        keys.append("m:" + x["email"].strip().lower())
    name = re.sub(r"[\s　・]|株式会社|（株）|\(株\)|㈱|有限会社|（有）|\(有\)|㈲|合同会社|合資会社|一般社団法人|一般財団法人|公益社団法人|公益財団法人|特定非営利活動法人|NPO法人|学校法人|社会福祉法人|医療法人社団|医療法人財団|医療法人", "", x.get("company") or "")
    if name:
        keys.append("n:" + name.lower())
    return keys

EIGYO_LINE_IN_MEMO = re.compile(r"^(?:個別の一文|一言|ひとこと)[：:]\s*(.+)$", re.M)

def eigyo_split_line(memo, line=""):
    """メモの中の「個別の一文：…」を取り出す。戻り値 (一文, 残りのメモ)"""
    memo = memo or ""
    m = EIGYO_LINE_IN_MEMO.search(memo)
    if not m:
        return (line or "").strip(), memo
    rest = (memo[:m.start()] + memo[m.end():].lstrip("\n")).strip("\n")
    return ((line or "").strip() or m.group(1).strip()), rest

def parse_eigyo_page(page):
    try:
        pr = page["properties"]
        d = {
            "id":       page["id"],
            "company":  "".join(t.get("plain_text", "") for t in (pr.get("会社名", {}).get("title") or [])).strip(),
            "industry": mail_rt_plain(pr, "業種").strip(),
            "site":     pr.get("サイトURL", {}).get("url") or "",
            "form":     pr.get("フォームURL", {}).get("url") or "",
            "email":    pr.get("メール", {}).get("email") or "",
            "person":   mail_rt_plain(pr, "担当者").strip(),
            "status":   (pr.get("ステータス", {}).get("select") or {}).get("name", "未着手"),
            "method":   (pr.get("送信方法", {}).get("select") or {}).get("name", ""),
            "sentAt":   (pr.get("送信日時", {}).get("date") or {}).get("start") or "",
            "template": mail_rt_plain(pr, "使ったテンプレート").strip(),
            "memo":     mail_rt_plain(pr, "メモ"),
            "line":     mail_rt_plain(pr, "個別の一文").strip(),
            "created":  page.get("created_time", ""),
        }
        d["line"], d["memo"] = eigyo_split_line(d["memo"], d["line"])
        return d
    except Exception:
        return None

def eigyo_props(x):
    x = dict(x)
    x["line"], x["memo"] = eigyo_split_line(x.get("memo"), x.get("line"))
    props = {
        "会社名":   {"title": [{"text": {"content": (x.get("company") or "（会社名なし）")[:MAIL_TPL_CHUNK]}}]},
        "業種":     {"rich_text": mail_rt_chunks((x.get("industry") or "").strip())},
        "サイトURL": {"url": eigyo_norm_url(x.get("site")) or None},
        "フォームURL": {"url": eigyo_norm_url(x.get("form")) or None},
        "メール":   {"email": (x.get("email") or "").strip() or None},
        "担当者":   {"rich_text": mail_rt_chunks((x.get("person") or "").strip())},
        "メモ":     {"rich_text": mail_rt_chunks(x.get("memo") or "")},
        "個別の一文": {"rich_text": mail_rt_chunks((x.get("line") or "").strip())},
    }
    st = (x.get("status") or "").strip()
    props["ステータス"] = {"select": {"name": st if st in EIGYO_STATUSES else "未着手"}}
    if "method" in x:
        props["送信方法"] = {"select": {"name": x["method"]}} if x.get("method") in ("フォーム", "メール") else {"select": None}
    if "sentAt" in x:
        props["送信日時"] = {"date": {"start": x["sentAt"]}} if x.get("sentAt") else {"date": None}
    if "template" in x:
        props["使ったテンプレート"] = {"rich_text": mail_rt_chunks(x.get("template") or "")}
    return props

def parse_eigyo_tpl_page(page):
    try:
        pr = page["properties"]
        return {
            "id":       page["id"],
            "name":     "".join(t.get("plain_text", "") for t in (pr.get("テンプレート名", {}).get("title") or [])).strip(),
            "industry": mail_rt_plain(pr, "業種").strip(),
            "subject":  mail_rt_plain(pr, "件名"),
            "body":     mail_rt_plain(pr, "本文"),
            "enabled":  pr.get("有効", {}).get("checkbox", True),
            "order":    pr.get("並び順", {}).get("number"),
        }
    except Exception:
        return None

def notion_query_all(db_id, body=None):
    items, cursor = [], None
    while True:
        b = dict(body or {}); b["page_size"] = 100
        if cursor:
            b["start_cursor"] = cursor
        result, err = notion_request("POST", f"/databases/{db_id}/query", b)
        if not result:
            return None, err
        items += result.get("results", [])
        if not result.get("has_more"):
            return items, None
        cursor = result.get("next_cursor")

EIGYO_OPTOUT = "今後このようなご案内が不要でしたら、お手数ですがこのメールに「不要」とご返信ください。以後お送りいたしません。"

def eigyo_mail_body(body):
    """特定電子メール法に沿って、送信者と配信停止の案内を必ず付ける"""
    body = (body or "").rstrip()
    tail = []
    # 本文（テンプレート）に無い送り主の情報だけを足す。住所は特定電子メール法で必須
    sig = []
    if EIGYO_SENDER_NAME and EIGYO_SENDER_NAME not in body: sig.append(EIGYO_SENDER_NAME)
    if EIGYO_SENDER_ADDRESS and EIGYO_SENDER_ADDRESS not in body: sig.append(f"所在地：{EIGYO_SENDER_ADDRESS}")
    if GMAIL_USER and GMAIL_USER not in body: sig.append(f"メール：{GMAIL_USER}")
    if EIGYO_SENDER_TEL and EIGYO_SENDER_TEL not in body: sig.append(f"電話：{EIGYO_SENDER_TEL}")
    if EIGYO_SENDER_URL and EIGYO_SENDER_URL not in body: sig.append(f"URL：{EIGYO_SENDER_URL}")
    if sig:
        tail.append("\n".join(sig))
    if "不要" not in body or "ご返信" not in body:
        tail.append("※" + EIGYO_OPTOUT)
    return body + ("\n\n――――――――――\n" + "\n".join(tail) if tail else "") + "\n"


# ---------- 営業メール：差し込み・テンプレ選択・送信の共通部分 ----------
def eigyo_fill(text, x):
    """画面の fill() と同じ差し込み。{{個別の一文}} が空なら、その段落ごと省く"""
    line = (x.get("line") or "").strip()
    paras = []
    for p in re.split(r"\n{2,}", text or ""):
        if "{{個別の一文}}" in p:
            if not line:
                continue
            p = p.replace("{{個別の一文}}", line)
        paras.append(p)
    text = "\n\n".join(paras)
    return (text.replace("{{会社名}}", x.get("company") or "貴社")
                .replace("{{担当者}}", x.get("person") or "ご担当者")
                .replace("{{業種}}", x.get("industry") or "")
                .replace("{{名前}}", EIGYO_SENDER_NAME).replace("{{メール}}", GMAIL_USER)
                .replace("{{電話}}", EIGYO_SENDER_TEL).replace("{{URL}}", EIGYO_SENDER_URL))

def eigyo_pick_template(tpls, x):
    """画面の tplsFor() と同じ優先順位：業種が一致 ＞ 部分一致 ＞ 業種なし（共通）"""
    on = [t for t in tpls if t.get("enabled")]
    ind = (x.get("industry") or "").strip()
    exact = [t for t in on if t["industry"] and ind and t["industry"] == ind]
    loose = [t for t in on if t["industry"] and ind and t["industry"] != ind and (t["industry"] in ind or ind in t["industry"])]
    common = [t for t in on if not t["industry"]]
    cand = exact + loose + common
    return cand[0] if cand else None

def _eigyo_count_today():
    today = eigyo_now_jst().date().isoformat()
    pages, _ = notion_query_all(EIGYO_LIST_DB_ID, {"filter": {"and": [
        {"property": "送信方法", "select": {"equals": "メール"}},
        {"property": "送信日時", "date": {"on_or_after": today}}]}})
    return len(pages or [])

def eigyo_send_core(to, subject, body, pid="", template_name="", force=False, dry=False, cur=None, all_items=None):
    """営業メールを1通送り、リストを送信済にする。戻り値 (ok, 結果dict または エラー文)"""
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header
    from email.utils import formataddr, make_msgid
    if not dry and not (GMAIL_USER and GMAIL_APP_PASSWORD):
        return False, "Railway に GMAIL_USER と GMAIL_APP_PASSWORD を設定してください"
    to = (to or "").strip(); subject = (subject or "").strip(); body = body or ""
    if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", to):
        return False, "宛先のメールアドレスが正しくありません"
    if not subject or not body.strip():
        return False, "件名と本文を入れてください"
    if "{{" in subject + body:
        return False, "差し込み（{{ }}）が残っています"
    if pid and cur is None:
        page, _ = notion_request("GET", f"/pages/{pid}")
        cur = parse_eigyo_page(page) if page else None
    if cur and cur["status"] in ("営業お断り", "対象外"):
        return False, f"この会社は「{cur['status']}」なので送れません"
    keys = set(eigyo_key({"email": to, **({k: cur[k] for k in ("site", "company")} if cur else {})}))
    if all_items is None:
        pages, _ = notion_query_all(EIGYO_LIST_DB_ID)
        all_items = [y for y in map(parse_eigyo_page, pages or []) if y]
    for y in all_items:
        if y["id"] == pid or not (keys & set(eigyo_key(y))):
            continue
        if y["status"] in ("営業お断り", "対象外"):
            return False, f"同じ会社の別の行（{y['company']}）が「{y['status']}」です"
        if y["status"] in ("送信済", "返信あり", "商談") and not force:
            when = (y["sentAt"] or "")[:10]
            return False, f"同じ会社（{y['company']}）に {when or '以前'} 送信済みです"
    if _eigyo_count_today() >= EIGYO_DAILY_LIMIT:
        return False, f"今日はもう {EIGYO_DAILY_LIMIT} 通送っています（1日の上限）"
    if not dry:
        msg = MIMEText(eigyo_mail_body(body), "plain", "utf-8")
        msg["Subject"] = Header(subject, "utf-8")
        msg["From"] = formataddr((str(Header(EIGYO_SENDER_NAME, "utf-8")), GMAIL_USER))
        msg["To"] = to
        msg["Message-ID"] = make_msgid(domain=GMAIL_USER.split("@")[-1])
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
                smtp.login(GMAIL_USER, GMAIL_APP_PASSWORD)
                smtp.sendmail(GMAIL_USER, [to], msg.as_string())
        except smtplib.SMTPAuthenticationError:
            return False, "AUTH:Gmail にログインできません。アプリパスワードを確認してください"
        except Exception as e:
            return False, f"送信できませんでした（{e.__class__.__name__}）"
    if pid and not dry:          # テスト送信ではリストを書き換えない
        notion_request("PATCH", f"/pages/{pid}", {"properties": {
            "ステータス": {"select": {"name": "送信済"}}, "送信方法": {"select": {"name": "メール"}},
            "送信日時": {"date": {"start": eigyo_now_jst().isoformat(timespec="seconds")}},
            "使ったテンプレート": {"rich_text": mail_rt_chunks(template_name or "")}}})
    print(f"  📧 営業メール送信{'（テスト・未送信）' if dry else ''}: {to}")
    return True, {"to": to, "subject": subject, "body": eigyo_mail_body(body), "dry": dry}

# ---------- 自動送信（バックグラウンドのスレッドが営業リストを順に送る） ----------
import threading as _threading
EIGYO_AUTO = {"enabled": EIGYO_AUTO_DEFAULT, "log": [], "last_sent_at": None, "last_check": None,
              "skipped": {}, "low_notified": "", "lock": _threading.Lock(), "remaining": None}

def _eigyo_auto_load():
    try:
        with open(EIGYO_AUTO_STATE_FILE, encoding="utf-8") as f:
            d = jsonlib.load(f)
        EIGYO_AUTO["enabled"] = bool(d.get("enabled", EIGYO_AUTO["enabled"]))
        EIGYO_AUTO["log"] = d.get("log", [])[-50:]
        EIGYO_AUTO["last_sent_at"] = d.get("last_sent_at")
        EIGYO_AUTO["low_notified"] = d.get("low_notified", "")
    except Exception:
        pass

def _eigyo_auto_save():
    try:
        with open(EIGYO_AUTO_STATE_FILE, "w", encoding="utf-8") as f:
            jsonlib.dump({"enabled": EIGYO_AUTO["enabled"], "log": EIGYO_AUTO["log"][-50:],
                          "last_sent_at": EIGYO_AUTO["last_sent_at"], "low_notified": EIGYO_AUTO["low_notified"]}, f, ensure_ascii=False)
    except Exception:
        pass

def _eigyo_auto_log(kind, text):
    EIGYO_AUTO["log"].append({"at": eigyo_now_jst().isoformat(timespec="seconds"), "kind": kind, "text": text})
    EIGYO_AUTO["log"] = EIGYO_AUTO["log"][-50:]
    print(f"  🤖 自動送信[{kind}] {text}")
    _eigyo_auto_save()

def _eigyo_parse_range(txt, lo, hi):
    try:
        a, b = [int(v) for v in str(txt).split("-")]
        return max(lo, a), min(hi, b)
    except Exception:
        return lo, hi

def eigyo_auto_in_window(now=None):
    """送ってよい時間帯か（曜日と時刻）"""
    now = now or eigyo_now_jst()
    d1, d2 = _eigyo_parse_range(EIGYO_AUTO_DAYS, 1, 7)
    h1, h2 = _eigyo_parse_range(EIGYO_AUTO_HOURS, 0, 24)
    return d1 <= now.isoweekday() <= d2 and h1 <= now.hour < h2

def eigyo_auto_candidates(all_items):
    """未着手でメールがあり、同じ会社に送信済・お断りが無いものを古い順に"""
    owner = {}
    for x in all_items:
        for k in eigyo_key(x):
            owner.setdefault(k, []).append(x)
    out = []
    for x in sorted(all_items, key=lambda y: y.get("created", "")):
        if x["status"] != "未着手" or not (x.get("email") or "").strip():
            continue
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", x["email"].strip()):
            continue
        blocked = False
        for k in eigyo_key(x):
            for y in owner.get(k, []):
                if y["id"] != x["id"] and y["status"] in ("送信済", "返信あり", "商談", "営業お断り", "対象外"):
                    blocked = True
        if not blocked and x["id"] not in EIGYO_AUTO["skipped"]:
            out.append(x)
    return out

def eigyo_auto_send_one(reason="auto"):
    """候補の先頭に1通送る。戻り値 (ok, メッセージ)"""
    with EIGYO_AUTO["lock"]:
        pages, err = notion_query_all(EIGYO_LIST_DB_ID)
        if pages is None:
            _eigyo_auto_log("エラー", f"リストが読めません: {err}"); return False, err
        all_items = [y for y in map(parse_eigyo_page, pages) if y]
        cands = eigyo_auto_candidates(all_items)
        EIGYO_AUTO["remaining"] = len(cands)
        tpages, _ = notion_query_all(EIGYO_TPL_DB_ID)
        tpls = [t for t in map(parse_eigyo_tpl_page, tpages or []) if t]
        if not cands:
            return False, "送れる未着手（メールあり）がありません"
        if not EIGYO_AUTO_DRY and not (GMAIL_USER and GMAIL_APP_PASSWORD):
            return False, "Railway に GMAIL_USER と GMAIL_APP_PASSWORD を設定してください"
        for x in cands:
            t = eigyo_pick_template(tpls, x)
            if not t:
                EIGYO_AUTO["skipped"][x["id"]] = "テンプレなし"
                _eigyo_auto_log("見送り", f"{x['company'] or x['email']}: 業種「{x['industry']}」に当たるテンプレートがありません"); continue
            sub, body = eigyo_fill(t["subject"], x), eigyo_fill(t["body"], x)
            ok, res = eigyo_send_core(x["email"], sub, body, pid=x["id"], template_name=t["name"],
                                      dry=EIGYO_AUTO_DRY, cur=x, all_items=all_items)
            if ok:
                EIGYO_AUTO["last_sent_at"] = eigyo_now_jst().isoformat(timespec="seconds")
                EIGYO_AUTO["remaining"] = len(cands) - 1
                EIGYO_AUTO["last_mail"] = res
                if EIGYO_AUTO_DRY:
                    EIGYO_AUTO["skipped"][x["id"]] = "テスト済"      # テストでは同じ会社を繰り返さない
                _eigyo_auto_log("送信" if not EIGYO_AUTO_DRY else "テスト", f"{x['company'] or ''} <{x['email']}> テンプレ「{t['name']}」（{reason}）")
                eigyo_auto_check_low(len(cands) - 1)
                return True, f"{x['company'] or x['email']} に{'送ったことにしました（テスト）' if EIGYO_AUTO_DRY else '送りました'}"
            msg = str(res)
            if msg.startswith("AUTH:"):
                _eigyo_auto_log("エラー", msg[5:] + "。自動送信を止めました"); EIGYO_AUTO["enabled"] = False; _eigyo_auto_save()
                return False, msg[5:]
            # 会社のせいではない失敗（設定不足・1日の上限・送信エラー）は、会社に印を付けずにここで止める
            if "GMAIL_USER" in msg or "上限" in msg or msg.startswith("送信できませんでした"):
                _eigyo_auto_log("エラー", msg + "（この会社は未着手のまま。次回に再挑戦）")
                return False, msg
            # この会社は送れない（宛先が不正・同じ会社に送信済み など）→ 送信失敗にして理由をメモに残し、次へ
            memo = (x.get("memo") or "").rstrip()
            notion_request("PATCH", f"/pages/{x['id']}", {"properties": {
                "ステータス": {"select": {"name": "送信失敗"}},
                "メモ": {"rich_text": mail_rt_chunks((f"自動送信エラー（{eigyo_now_jst().date().isoformat()}）：{msg}\n" + memo).strip())}}})
            _eigyo_auto_log("失敗", f"{x['company'] or x['email']}: {msg}")
            _time.sleep(0.4)
        return False, "送れる候補がありませんでした"

def eigyo_auto_check_low(remaining):
    """残りが少なくなったら引き継ぎ（やること）に「営業リストの補充」を1日1回だけ書く"""
    today = eigyo_now_jst().date().isoformat()
    if remaining is None or remaining >= EIGYO_LOW_WATER or EIGYO_AUTO["low_notified"] == today:
        return
    props = {"内容": {"title": [{"text": {"content": f"営業リストの補充（メール営業の未着手が残り {remaining} 件）"}}]},
             "種別": {"select": {"name": "todo"}}, "ステータス": {"select": {"name": "active"}},
             "作成日": {"date": {"start": today}}, "日付": {"date": {"start": today}}}
    r, _ = notion_request("POST", "/pages", {"parent": {"database_id": HANDOVER_DB_ID}, "properties": props})
    if r:
        EIGYO_AUTO["low_notified"] = today
        _eigyo_auto_log("補充", f"残り {remaining} 件。引き継ぎに「営業リストの補充」を書きました")

def eigyo_auto_status():
    d1, d2 = _eigyo_parse_range(EIGYO_AUTO_DAYS, 1, 7); h1, h2 = _eigyo_parse_range(EIGYO_AUTO_HOURS, 0, 24)
    return {"enabled": EIGYO_AUTO["enabled"], "dry": EIGYO_AUTO_DRY, "configured": bool(GMAIL_USER and GMAIL_APP_PASSWORD),
            "user": GMAIL_USER, "sender": EIGYO_SENDER_NAME, "address": EIGYO_SENDER_ADDRESS,
            "window": {"days": [d1, d2], "hours": [h1, h2]}, "inWindow": eigyo_auto_in_window(),
            "perDay": EIGYO_AUTO_PER_DAY, "limit": EIGYO_DAILY_LIMIT, "gapSec": EIGYO_AUTO_GAP_SEC, "lowWater": EIGYO_LOW_WATER,
            "lastSentAt": EIGYO_AUTO["last_sent_at"], "lastCheck": EIGYO_AUTO["last_check"],
            "remaining": EIGYO_AUTO["remaining"], "log": EIGYO_AUTO["log"][-30:][::-1]}

def eigyo_auto_loop():
    _eigyo_auto_load()
    _time.sleep(20)
    while True:
        try:
            EIGYO_AUTO["last_check"] = eigyo_now_jst().isoformat(timespec="seconds")
            if EIGYO_AUTO["enabled"] and eigyo_auto_in_window():
                last = EIGYO_AUTO["last_sent_at"]
                due = True
                if last:
                    try:
                        due = (eigyo_now_jst() - datetime.datetime.fromisoformat(last)).total_seconds() >= EIGYO_AUTO_GAP_SEC
                    except Exception:
                        due = True
                if due and _eigyo_count_today() < min(EIGYO_AUTO_PER_DAY, EIGYO_DAILY_LIMIT):
                    eigyo_auto_send_one("自動")
        except Exception as e:
            _eigyo_auto_log("エラー", f"{e.__class__.__name__}: {e}")
        _time.sleep(60)

def notion_request(method, path, body=None):
    url = NOTION_API + path
    data = jsonlib.dumps(body).encode() if body else None
    req = Request(url, data=data, headers=HEADERS, method=method)
    try:
        with urlopen(req, timeout=15) as res:
            return jsonlib.loads(res.read()), None
    except HTTPError as e:
        err_body = e.read().decode(errors='replace')
        print(f"  ❌ Notion HTTP {e.code}: {err_body[:300]}")
        return None, f"Notion API {e.code}: {err_body[:200]}"
    except URLError as e:
        print(f"  ❌ Notion URLError: {e.reason}")
        return None, f"Notion接続エラー: {e.reason}"
    except Exception as e:
        print(f"  ❌ Notion例外: {e}")
        return None, f"例外: {e}"

# ========== 請求書 Notion ヘルパー ==========

def parse_notion_invoice_page(page):
    """Notion請求書ページをdictに変換"""
    try:
        props = page["properties"]
        def rt(key, default=""):
            lst = (props.get(key, {}).get("rich_text") or [])
            return lst[0].get("plain_text", default) if lst else default
        def title_prop():
            t = (props.get("タイトル", {}).get("title") or [])
            return t[0].get("plain_text", "") if t else ""
        def sel(key, default=""):
            return (props.get(key, {}).get("select") or {}).get("name", default)
        def num(key, default=0):
            v = props.get(key, {}).get("number")
            return v if v is not None else default
        def chk(key):
            return props.get(key, {}).get("checkbox", False)
        def date_prop(key):
            return (props.get(key, {}).get("date") or {}).get("start", "")
        return {
            "id":             title_prop(),
            "_notion_id":     page["id"],
            "received_at":    date_prop("受信日時"),
            "worker_name":    rt("ワーカー名"),
            "billing_month":  rt("請求月"),
            "amount":         num("金額"),
            "status":         sel("ステータス", "new"),
            "bank_name":      rt("金融機関名"),
            "branch":         rt("支店"),
            "account_type":   sel("口座種別", "普通"),
            "account_number": rt("口座番号"),
            "account_holder": rt("口座名義"),
            "remark":         rt("備考"),
            "has_pdf":        chk("PDFあり"),
            "worker_id":      rt("ワーカーID"),
        }
    except Exception as e:
        print(f"  ⚠️ parse_notion_invoice_page: {e}")
        return None

def invoice_to_notion_props(inv):
    """請求書dictをNotion propertiesに変換"""
    received_at = inv.get("received_at", "")
    date_str = received_at[:10] if received_at else datetime.date.today().isoformat()
    return {
        "タイトル":   {"title":     [{"text": {"content": inv["id"]}}]},
        "ワーカー名": {"rich_text": [{"text": {"content": inv.get("worker_name", "")}}]},
        "請求月":     {"rich_text": [{"text": {"content": inv.get("billing_month", "")}}]},
        "金額":       {"number":    inv.get("amount", 0)},
        "ステータス": {"select":    {"name": inv.get("status", "new")}},
        "金融機関名": {"rich_text": [{"text": {"content": inv.get("bank_name", "")}}]},
        "支店":       {"rich_text": [{"text": {"content": inv.get("branch", "")}}]},
        "口座種別":   {"select":    {"name": inv.get("account_type", "普通")}},
        "口座番号":   {"rich_text": [{"text": {"content": inv.get("account_number", "")}}]},
        "口座名義":   {"rich_text": [{"text": {"content": inv.get("account_holder", "")}}]},
        "備考":       {"rich_text": [{"text": {"content": inv.get("remark", "")}}]},
        "受信日時":   {"date":      {"start": date_str}},
        "PDFあり":    {"checkbox":  inv.get("has_pdf", False)},
        "ワーカーID": {"rich_text": [{"text": {"content": inv.get("worker_id", "") or ""}}]},
    }

def get_invoices_from_notion():
    """Notionから請求書一覧を取得（新しい順）"""
    if not INVOICE_DB_ID:
        return []
    results = []
    body = {
        "sorts": [{"timestamp": "created_time", "direction": "descending"}],
        "page_size": 100,
    }
    cursor = None
    while True:
        if cursor: body["start_cursor"] = cursor
        result, _ = notion_request("POST", f"/databases/{INVOICE_DB_ID}/query", body)
        if not result: break
        for page in result.get("results", []):
            inv = parse_notion_invoice_page(page)
            if inv: results.append(inv)
        if not result.get("has_more"): break
        cursor = result.get("next_cursor")
    return results

def find_invoice_notion_id(inv_id):
    """請求書IDからNotionページIDを検索"""
    if not INVOICE_DB_ID: return None
    body = {
        "filter": {"property": "タイトル", "title": {"equals": inv_id}},
        "page_size": 1,
    }
    result, _ = notion_request("POST", f"/databases/{INVOICE_DB_ID}/query", body)
    if result and result.get("results"):
        return result["results"][0]["id"]
    return None

# ========== ワーカー Notion ヘルパー ==========

def parse_notion_worker_page(page):
    """Notionワーカーページをdictに変換"""
    try:
        props = page["properties"]
        def rt(key, default=""):
            lst = (props.get(key, {}).get("rich_text") or [])
            return lst[0].get("plain_text", default) if lst else default
        def title_prop():
            t = (props.get("タイトル", {}).get("title") or [])
            return t[0].get("plain_text", "") if t else ""
        def sel(key, default=""):
            return (props.get(key, {}).get("select") or {}).get("name", default)
        return {
            "worker_id":  title_prop(),
            "_notion_id": page["id"],
            "name":       rt("名前"),
            "pw_hash":    rt("パスワードハッシュ"),
            "token":      rt("トークン"),
            "bankName":   rt("金融機関名"),
            "bankBranch": rt("支店"),
            "bankNo":     rt("口座番号"),
            "bankHolder": rt("口座名義"),
            "acctType":   sel("口座種別", "普通"),
            "note":       rt("備考メモ"),
        }
    except Exception as e:
        print(f"  ⚠️ parse_notion_worker_page: {e}")
        return None

def worker_to_notion_props(w):
    """ワーカーdictをNotion propertiesに変換"""
    return {
        "タイトル":           {"title":     [{"text": {"content": w.get("worker_id", "")}}]},
        "名前":               {"rich_text": [{"text": {"content": w.get("name", "")}}]},
        "パスワードハッシュ": {"rich_text": [{"text": {"content": w.get("pw_hash", "")}}]},
        "トークン":           {"rich_text": [{"text": {"content": w.get("token", "")}}]},
        "金融機関名":         {"rich_text": [{"text": {"content": w.get("bankName", "")}}]},
        "支店":               {"rich_text": [{"text": {"content": w.get("bankBranch", "")}}]},
        "口座番号":           {"rich_text": [{"text": {"content": w.get("bankNo", "")}}]},
        "口座名義":           {"rich_text": [{"text": {"content": w.get("bankHolder", "")}}]},
        "口座種別":           {"select":    {"name": w.get("acctType", "普通")}},
        "備考メモ":           {"rich_text": [{"text": {"content": w.get("note", "")}}]},
    }

def get_worker_by_id_notion(worker_id):
    """ワーカーIDでNotionからワーカーを検索"""
    if not WORKER_DB_ID: return None
    body = {
        "filter": {"property": "タイトル", "title": {"equals": worker_id}},
        "page_size": 1,
    }
    result, _ = notion_request("POST", f"/databases/{WORKER_DB_ID}/query", body)
    if result and result.get("results"):
        return parse_notion_worker_page(result["results"][0])
    return None

def get_worker_by_token_notion(token):
    """トークンでNotionからワーカーを検索"""
    if not WORKER_DB_ID or not token: return None
    body = {
        "filter": {"property": "トークン", "rich_text": {"equals": token}},
        "page_size": 1,
    }
    result, _ = notion_request("POST", f"/databases/{WORKER_DB_ID}/query", body)
    if result and result.get("results"):
        return parse_notion_worker_page(result["results"][0])
    return None

# ========== 顧客番号ユーティリティ ==========
def get_all_customer_nos():
    """お客様DBから全番号を取得"""
    nos = []
    cursor = None
    while True:
        body = {"page_size": 100}
        if cursor:
            body["start_cursor"] = cursor
        result, _ = notion_request("POST", f"/databases/{CUSTOMER_DB_ID}/query", body)
        if not result:
            break
        for page in result.get("results", []):
            try:
                rt = page["properties"]["お客様No."]["rich_text"]
                no = rt[0]["plain_text"].strip() if rt else ""
                if no:
                    nos.append(no)
            except Exception:
                pass
        if not result.get("has_more"):
            break
        cursor = result.get("next_cursor")
    return nos

def next_in_range(nos, lo, hi):
    used = set()
    for no in nos:
        try:
            n = int(no)
            if lo <= n <= hi:
                used.add(n)
        except Exception:
            pass
    for n in range(lo, hi + 2):
        if n not in used:
            return str(n)

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"  {args[0]} {args[1]}")

    def send_json(self, code, obj):
        body = jsonlib.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def send_html(self, filename):
        filepath = os.path.join(SCRIPT_DIR, filename)
        if not os.path.exists(filepath):
            self.send_response(404)
            self.end_headers()
            return
        with open(filepath, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def send_static(self, filename, content_type):
        """JS/CSSなどの静的ファイルを返す（HTML以外用）"""
        filepath = os.path.join(SCRIPT_DIR, filename)
        if not os.path.exists(filepath):
            self.send_response(404)
            self.end_headers()
            return
        with open(filepath, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", len(body))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        from urllib.parse import unquote
        path = unquote(self.path.split('?')[0])  # デコード＆クエリ除去

        if path in ("/", "/index.html"):
            self.send_html("index.html")
        elif path == "/案件登録ツール.html":
            self.send_html("案件登録ツール.html")
        elif path == "/顧客登録ツール.html":
            self.send_html("顧客登録ツール.html")
        elif path == "/請求書ツール.html":
            self.send_html("請求書ツール.html")
        elif path == "/顧客情報ツール.html":
            self.send_html("顧客情報ツール.html")
        elif path == "/引き継ぎ一覧.html":
            self.send_html("引き継ぎ一覧.html")
        elif path == "/引き継ぎ完了済み.html":
            self.send_html("引き継ぎ完了済み.html")
        elif path == "/日次スケジュール.html":
            self.send_html("日次スケジュール.html")
        elif path == "/案件履歴ツール.html":
            self.send_html("案件履歴ツール.html")
        elif path == "/請求書管理ツール.html":
            self.send_html("請求書管理ツール.html")
        elif path == "/ワーカー請求書.html":
            self.send_html("ワーカー請求書.html")
        elif path == "/メール定型文ツール.html":
            self.send_html("メール定型文ツール.html")
        elif path == "/日報ツール.html":
            self.send_html("日報ツール.html")
        elif path == "/営業ツール.html":
            self.send_html("営業ツール.html")
        elif path == "/api/invoices":
            invoices = get_invoices_from_notion()
            self.send_json(200, invoices)
        elif path.startswith("/api/invoice-pdf/"):
            inv_id = path.split("/")[-1]
            pdf_path = os.path.join(INVOICE_PDF_DIR, inv_id + ".pdf")
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "application/pdf")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(data)
            else:
                self.send_json(404, {"error": "PDF not found"})
        elif path == "/api/health":
            self.send_json(200, {
                "status": "ok",
                "message": "サーバー起動中",
                "invoice_db_set": bool(INVOICE_DB_ID),
                "worker_db_set":  bool(WORKER_DB_ID),
                "notion_token_set": bool(NOTION_TOKEN),
            })
        elif path == "/api/myip":
            ip = self.headers.get("X-Forwarded-For", self.client_address[0]).split(",")[0].strip()
            self.send_json(200, {"ip": ip})
        elif path == "/api/next-customer-no":
            nos = get_all_customer_nos()
            self.send_json(200, {
                "ok": True,
                "next_200s":     next_in_range(nos, 200, 299),   # 企業直接・裁量低め/担当者経由
                "next_300s":     next_in_range(nos, 300, 349),   # 直営業・重要/決裁権窓口
                "next_referral": next_in_range(nos, 350, 399),   # 紹介・つながり案件
                "next_3000s":    next_in_range(nos, 3000, 3999), # 新規チャンネル・裁量高
                "next_regular":  next_in_range(nos, 1, 99),      # 通常（固有値・再利用不可）
            })
        elif path == "/mail_template.js":
            self.send_static("mail_template.js", "application/javascript; charset=utf-8")
        elif path == "/customer_link.js":
            self.send_static("customer_link.js", "application/javascript; charset=utf-8")
        elif path == "/api/customer-links":
            # 全ページ共通：お客様No→Notion顧客ページURL（?refresh=1 で再構築）
            links = get_customer_link_map(force=("refresh=" in self.path))
            self.send_json(200, {"ok": True, "count": len(links), "links": links})
        elif path == "/api/customers":
            self.handle_get_customers()
        elif path == "/api/customers-all":
            self.handle_get_all_customers()
        elif path == "/api/mail-templates":
            self.handle_mail_templates_list()
        elif path == "/api/nippo":
            self.handle_nippo_list()
        elif path.startswith("/api/invoice-data"):
            self.handle_get_invoice_data()
        elif path.startswith("/api/calendar/day"):
            self.handle_get_calendar_day()
        elif path.startswith("/api/calendar"):
            self.handle_get_calendar()
        elif path == "/api/handover/all":
            self.handle_get_handover_all()
        elif path == "/api/handover/done-list":
            self.handle_get_handover_done_list()
        elif path.startswith("/api/cases/list"):
            self.handle_get_cases_list()
        elif path.startswith("/api/daily-schedule/dates"):
            self.handle_get_schedule_dates()
        elif path.startswith("/api/daily-schedule"):
            self.handle_get_daily_schedule()
        elif path.startswith("/api/handover"):
            self.handle_get_handover()
        else:
            self.send_json(404, {"error": "Not found"})

    def do_DELETE(self):
        path = unquote(self.path.split('?')[0])
        if path.startswith("/api/invoice/") and path.count("/") == 3:
            inv_id = path.split("/")[-1]
            notion_page_id = find_invoice_notion_id(inv_id)
            if notion_page_id:
                notion_request("PATCH", f"/pages/{notion_page_id}", {"archived": True})
                print(f"  🗑️  請求書削除（アーカイブ）: {inv_id}")
            self.send_json(200, {"ok": True})
        else:
            self.send_json(404, {"error": "Not found"})

    def do_PATCH(self):
        path = unquote(self.path.split('?')[0])
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        try:
            data = jsonlib.loads(raw)
        except:
            self.send_json(400, {"error": "Invalid JSON"}); return
        # /api/invoice/{id}/status
        parts = path.strip("/").split("/")
        if len(parts) == 4 and parts[0] == "api" and parts[1] == "invoice" and parts[3] == "status":
            inv_id = parts[2]
            new_status = data.get("status", "new")
            notion_page_id = find_invoice_notion_id(inv_id)
            if notion_page_id:
                notion_request("PATCH", f"/pages/{notion_page_id}", {
                    "properties": {"ステータス": {"select": {"name": new_status}}}
                })
                print(f"  ✅ 請求書ステータス更新: {inv_id} → {new_status}")
            self.send_json(200, {"ok": True})
        else:
            self.send_json(404, {"error": "Not found"})

    def do_POST(self):
        path = unquote(self.path.split('?')[0])
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        content_type = self.headers.get("Content-Type", "")

        # ── multipart/form-data（請求書PDF送信）──────────
        if "multipart/form-data" in content_type:
            if path == "/api/submit-invoice":
                self.handle_submit_invoice(raw, content_type)
            else:
                self.send_json(404, {"error": "Not found"})
            return

        # ── JSON系エンドポイント ──────────────────────────
        try:
            data = jsonlib.loads(raw) if raw else {}
        except Exception:
            self.send_json(400, {"error": "Invalid JSON"})
            return

        if path == "/api/admin/login":
            pw = data.get("password", "")
            pw_token = _make_admin_token(pw)
            if ADMIN_TOKEN and pw_token == ADMIN_TOKEN:
                self.send_json(200, {"ok": True, "token": ADMIN_TOKEN})
            else:
                self.send_json(401, {"ok": False, "error": "パスワードが違います"})
            return
        elif path == "/api/admin/check":
            token = data.get("token", "")
            ok = bool(ADMIN_TOKEN) and token == ADMIN_TOKEN
            self.send_json(200, {"ok": ok})
            return
        if path == "/api/worker/login":
            self.handle_worker_login(data)
            return
        elif path == "/api/worker/register":
            self.handle_worker_register(data)
            return
        elif path.startswith("/api/register-worker-from-invoice/"):
            inv_id = path.split("/")[-1]
            self.handle_register_worker_from_invoice(inv_id, data)
            return
        if path.startswith("/api/eigyo/"):
            self.handle_eigyo(path, data)
            return
        if self.path == "/api/register":
            self.handle_register(data)
        elif self.path == "/api/update-customer":
            self.handle_update_customer(data)
        elif self.path == "/api/register-customer":
            self.handle_register_customer(data)
        elif self.path == "/api/record-invoice":
            self.handle_record_invoice(data)
        elif self.path == "/api/generate-invoice":
            self.handle_generate_invoice(data)
        elif self.path == "/api/calendar/add":
            self.handle_add_schedule(data)
        elif self.path == "/api/calendar/delete":
            self.handle_delete_schedule(data)
        elif self.path == "/api/customers/bulk-update":
            self.handle_bulk_update_customers(data)
        elif self.path == "/api/customers/bulk-archive":
            self.handle_bulk_archive_customers(data)
        elif self.path == "/api/mail-templates/save":
            self.handle_mail_template_save(data)
        elif self.path == "/api/mail-templates/delete":
            self.handle_mail_template_delete(data)
        elif self.path == "/api/nippo/save":
            self.handle_nippo_save(data)
        elif self.path == "/api/nippo/delete":
            self.handle_nippo_delete(data)
        elif self.path == "/api/customers/invoice-name":
            self.handle_set_invoice_name(data)
        elif self.path == "/api/handover/add":
            self.handle_handover_add(data)
        elif self.path == "/api/handover/done":
            self.handle_handover_done(data)
        elif self.path == "/api/handover/carry":
            self.handle_handover_carry(data)
        elif self.path == "/api/handover/extend":
            self.handle_handover_extend(data)
        elif self.path == "/api/handover/delete":
            self.handle_handover_delete(data)
        elif self.path == "/api/handover/update-date":
            self.handle_handover_update_date(data)
        elif self.path == "/api/handover/update-content":
            self.handle_handover_update_content(data)
        elif self.path == "/api/handover/restore":
            self.handle_handover_restore(data)
        elif self.path == "/api/daily-schedule/save":
            self.handle_save_daily_block(data)
        elif self.path == "/api/daily-schedule/delete":
            self.handle_delete_daily_block(data)
        elif self.path == "/api/notion-event/move":
            self.handle_move_notion_event(data)
        elif self.path == "/api/cases/update":
            self.handle_update_case(data)
        elif self.path == "/api/daily-schedule/gantt":
            self.handle_gantt_excel(data)
        elif self.path == "/api/invoice/template":
            self.handle_invoice_template(data)
        elif self.path == "/api/invoice/numbers":
            self.handle_invoice_numbers(data)
        else:
            self.send_json(404, {"error": "Not found"})

    # ── 請求書送信（multipart → Notion）──────────────────────────────
    def handle_submit_invoice(self, raw_body, content_type):
        try:
            fields, files = parse_multipart(raw_body, content_type)
        except Exception as e:
            self.send_json(400, {"ok": False, "error": f"パースエラー: {e}"}); return

        if not INVOICE_DB_ID:
            self.send_json(500, {"ok": False,
                "error": "INVOICE_DB_ID が設定されていません。Railway の Variables タブで設定してください"}); return

        inv_id = "inv-" + str(uuid.uuid4())[:8]
        now = datetime.datetime.now().isoformat()

        # ワーカートークンで名前を補完
        token = fields.get("workerToken", "")
        worker_name = fields.get("name", "").strip()
        if not worker_name and token:
            w = get_worker_by_token_notion(token)
            if w: worker_name = w.get("name", "")

        invoice = {
            "id":             inv_id,
            "received_at":    now,
            "worker_name":    worker_name,
            "billing_month":  fields.get("billingMonth", ""),
            "amount":         int(fields.get("amount", 0) or 0),
            "status":         "new",
            "bank_name":      fields.get("bankName", ""),
            "branch":         fields.get("bankBranch", ""),
            "account_type":   fields.get("acctType", "普通"),
            "account_number": fields.get("bankNo", ""),
            "account_holder": fields.get("bankHolder", ""),
            "remark":         fields.get("remarks", ""),
            "has_pdf":        "pdf" in files,
            "worker_id":      "",
        }

        # PDF一時保存（Railwayでは再デプロイ時に消えるが、請求書データはNotionに永続保存）
        if "pdf" in files:
            _, pdf_bytes = files["pdf"]
            pdf_path = os.path.join(INVOICE_PDF_DIR, inv_id + ".pdf")
            try:
                with open(pdf_path, "wb") as f:
                    f.write(pdf_bytes)
            except Exception as e:
                print(f"  ⚠️ PDF一時保存失敗: {e}")

        # Notionに保存
        result, err = notion_request("POST", "/pages", {
            "parent": {"database_id": INVOICE_DB_ID},
            "properties": invoice_to_notion_props(invoice),
        })
        if not result:
            self.send_json(500, {"ok": False, "error": f"Notion保存失敗: {err}"}); return

        print(f"  ✅ 請求書Notion保存: {inv_id} ({worker_name} {invoice['billing_month']} ¥{invoice['amount']:,})")
        self.send_json(200, {"ok": True, "invoiceId": inv_id})

    # ── ワーカー認証（Notion）────────────────────────────────────────
    def handle_worker_login(self, data):
        worker_id = data.get("workerId", "").strip()
        password  = data.get("password", "")
        if not WORKER_DB_ID:
            self.send_json(500, {"ok": False, "error": "WORKER_DB_ID が設定されていません"}); return
        w = get_worker_by_id_notion(worker_id)
        if not w:
            self.send_json(401, {"ok": False, "error": "IDが見つかりません"}); return
        pw_hash = hashlib.sha256(password.encode()).hexdigest()
        if w.get("pw_hash") != pw_hash:
            self.send_json(401, {"ok": False, "error": "パスワードが違います"}); return
        # トークンがなければ新規発行してNotionに保存
        token = w.get("token") or str(uuid.uuid4())
        if not w.get("token"):
            notion_id = w.get("_notion_id")
            if notion_id:
                notion_request("PATCH", f"/pages/{notion_id}", {
                    "properties": {"トークン": {"rich_text": [{"text": {"content": token}}]}}
                })
        profile = {k: w.get(k) for k in ["name","bankName","bankBranch","bankNo","bankHolder","acctType"]}
        self.send_json(200, {"ok": True, "token": token, "profile": profile})

    def handle_worker_register(self, data):
        worker_id = data.get("workerId", "").strip()
        password  = data.get("password", "")
        name      = data.get("name", "").strip()
        if not worker_id or not password or not name:
            self.send_json(400, {"ok": False, "error": "必須項目が不足しています"}); return
        if len(password) < 8:
            self.send_json(400, {"ok": False, "error": "パスワードは8文字以上"}); return
        if not WORKER_DB_ID:
            self.send_json(500, {"ok": False, "error": "WORKER_DB_ID が設定されていません"}); return
        existing = get_worker_by_id_notion(worker_id)
        if existing:
            self.send_json(409, {"ok": False, "error": "このIDは既に使われています"}); return
        token   = str(uuid.uuid4())
        pw_hash = hashlib.sha256(password.encode()).hexdigest()
        new_w = {
            "worker_id": worker_id, "name": name, "pw_hash": pw_hash, "token": token,
            "bankName":"","bankBranch":"","bankNo":"","bankHolder":"","acctType":"普通","note":"",
        }
        result, err = notion_request("POST", "/pages", {
            "parent": {"database_id": WORKER_DB_ID},
            "properties": worker_to_notion_props(new_w),
        })
        if not result:
            self.send_json(500, {"ok": False, "error": f"Notion保存失敗: {err}"}); return
        profile = {k: new_w.get(k) for k in ["name","bankName","bankBranch","bankNo","bankHolder","acctType"]}
        self.send_json(200, {"ok": True, "token": token, "profile": profile})

    # ── ワーカー登録（請求書から）──────────────────────────
    def handle_register_worker_from_invoice(self, inv_id, data):
        worker_id = data.get("worker_id", "").strip()
        note      = data.get("note", "")
        if not worker_id:
            self.send_json(400, {"ok": False, "error": "ワーカーIDが必要"}); return

        # 請求書のNotionページIDを検索
        notion_page_id = find_invoice_notion_id(inv_id)
        if not notion_page_id:
            self.send_json(404, {"ok": False, "error": "請求書が見つかりません"}); return

        # 請求書のワーカーIDをNotionで更新
        notion_request("PATCH", f"/pages/{notion_page_id}", {
            "properties": {"ワーカーID": {"rich_text": [{"text": {"content": worker_id}}]}}
        })

        # 請求書データを取得してワーカー情報を補完
        inv_result, _ = notion_request("GET", f"/pages/{notion_page_id}")
        inv = parse_notion_invoice_page(inv_result) if inv_result else {}

        # ワーカーが未登録なら作成
        if WORKER_DB_ID:
            existing = get_worker_by_id_notion(worker_id)
            if not existing:
                new_w = {
                    "worker_id": worker_id,
                    "name":        inv.get("worker_name", ""),
                    "pw_hash":     "", "token": "",
                    "bankName":    inv.get("bank_name", ""),
                    "bankBranch":  inv.get("branch", ""),
                    "bankNo":      inv.get("account_number", ""),
                    "bankHolder":  inv.get("account_holder", ""),
                    "acctType":    inv.get("account_type", "普通"),
                    "note":        note,
                }
                notion_request("POST", "/pages", {
                    "parent": {"database_id": WORKER_DB_ID},
                    "properties": worker_to_notion_props(new_w),
                })
                print(f"  👤 ワーカー登録: {worker_id}")
        self.send_json(200, {"ok": True})

    def send_file(self, filepath, filename, content_type):
        from urllib.parse import quote
        with open(filepath, 'rb') as f:
            data = f.read()
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{quote(filename)}")
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Expose-Headers', 'Content-Disposition')
        self.end_headers()
        self.wfile.write(data)

    def handle_generate_invoice(self, data):
        if not OPENPYXL_OK:
            self.send_json(500, {'ok': False, 'error': 'openpyxl未インストール。pip install openpyxl --break-system-packages を実行してください'})
            return
        try:
            self._generate_invoice_inner(data)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_json(500, {'ok': False, 'error': f'請求書生成エラー: {e}'})

    # 請求書雛形.xlsx の構造（この定数を変えるだけで雛形差し替えに追従できる）
    INV_TEMPLATE      = "請求書雛形.xlsx"
    # 金額の表示形式：小数点なし・3桁カンマ区切り（マイナスは ¥-1,234）
    INV_MONEY_FMT     = '"¥"#,##0;"¥"\\-#,##0'
    INV_ROW_NO        = 1    # B1: 請求書 No.
    INV_ROW_CUSTOMER  = 2    # B2: 宛名
    INV_ROW_DATE      = 8    # B8: 請求日 / C8: 見出し「内容」
    INV_ROW_DUE       = 9    # B9: 支払い期日 / C9: 内容
    INV_DATA_START    = 13   # 明細の先頭行（12行目は見出し）
    INV_DATA_ROWS     = 5    # 雛形が最初から持つ明細行数（13〜17）
    INV_ROW_SUMHEAD   = 18   # 「説明 / 金額」
    INV_ROW_SUBTOTAL  = 19   # 小計
    INV_ROW_TAX       = 20   # 税率
    INV_ROW_OTHER     = 21   # その他
    INV_ROW_TOTAL     = 22   # 集計

    def lookup_invoice_name(self, customer_no):
        """お客様Noから請求書の宛名を引く。請求書宛名 → クライアント名 の順。"""
        body = {"filter": {"property": "お客様No.", "rich_text": {"equals": customer_no}},
                "page_size": 1}
        res, _ = notion_request("POST", f"/databases/{CUSTOMER_DB_ID}/query", body)
        if not (res and res.get("results")):
            return ""
        props = res["results"][0].get("properties", {})
        for key in ("請求書宛名", "クライアント名"):
            rt = props.get(key, {}).get("rich_text") or []
            v = "".join(t.get("plain_text", "") for t in rt).strip()
            if v:
                return v
        return ""

    def _generate_invoice_inner(self, data):
        """請求書雛形.xlsx をテンプレートとして読み込み、値だけ差し込んで出力する。

        ゼロから組み立てるとフォント・テーマ色・テーブルスタイル・印刷設定が
        再現できないため、雛形そのものを土台にする方式に変更（2026-09-01）。
        """
        import copy as _copy
        import calendar as _calendar
        from openpyxl import load_workbook

        customer_no  = data.get('customerNo', '')
        invoice_date = data.get('invoiceDate', '')
        cases        = data.get('cases', [])

        # ---- 宛名 ----
        # 「請求書宛名」があればそれを使う。請求書だけ法人名や屋号で出したい、
        # という依頼に対応するための項目。空ならこれまで通りクライアント名。
        customer_name = (data.get('customerName') or '').strip()
        if not customer_name:
            customer_name = self.lookup_invoice_name(customer_no)
        if not customer_name:
            customer_name = customer_no
        # 既に敬称が付いている宛名（御中・殿など）はそのまま尊重する
        if not customer_name.endswith(("様", "御中", "殿")):
            customer_name += "様"

        # ---- 日付を雛形と同じ yyyy/m/d に正規化し、支払い期日（翌月末）を計算 ----
        def parse_date(text):
            text = (text or '').strip()
            for pat in (r'(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})',):
                m = re.search(pat, text)
                if m:
                    return int(m.group(1)), int(m.group(2)), int(m.group(3))
            return None

        ymd = parse_date(invoice_date)
        if ymd:
            y, mo, d = ymd
            invoice_str = f'{y}/{mo}/{d}'
            ny, nm = (y, mo + 1) if mo < 12 else (y + 1, 1)
            due_str = f'{ny}/{nm}/{_calendar.monthrange(ny, nm)[1]}'
        else:
            invoice_str = invoice_date
            due_str = ''

        # ---- 雛形を読み込む ----
        tpl = os.path.join(SCRIPT_DIR, self.INV_TEMPLATE)
        if not os.path.exists(tpl):
            self.send_json(500, {'ok': False,
                                 'error': f'{self.INV_TEMPLATE} が見つかりません（サーバー直下に配置してください）'})
            return
        wb = load_workbook(tpl)
        ws = wb.active

        # ---- 明細の行数を実件数に合わせる ----
        n = max(1, len(cases))
        extra = max(0, n - self.INV_DATA_ROWS)
        if extra:
            ins_at = self.INV_DATA_START + self.INV_DATA_ROWS

            # openpyxl の insert_rows は結合セル範囲を追従させない。
            # そのままだと「その他」「集計」の金額セルが B列との結合に飲み込まれ、
            # 値が入っていても Excel 上に出なくなる（実際に発生したバグ）。
            # ★順序が重要：挿入前に解除 → 挿入 → 位置をずらして再結合。
            #   挿入後に unmerge すると openpyxl が該当セルを空セルで作り直し、
            #   ずれ込んできた金額が消えてしまう。
            to_shift = [(r.min_row, r.min_col, r.max_row, r.max_col)
                        for r in list(ws.merged_cells.ranges) if r.min_row >= ins_at]
            for r0, c0, r1, c1 in to_shift:
                ws.unmerge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)

            ws.insert_rows(ins_at, extra)

            for r0, c0, r1, c1 in to_shift:
                ws.merge_cells(start_row=r0 + extra, start_column=c0,
                               end_row=r1 + extra, end_column=c1)

            # 追加行の書式・行高を雛形の明細行から複製
            src = self.INV_DATA_START
            for i in range(extra):
                dst = ins_at + i
                ws.row_dimensions[dst].height = ws.row_dimensions[src].height
                for col in ('B', 'C'):
                    sc, dc = ws[f'{col}{src}'], ws[f'{col}{dst}']
                    dc._style = _copy.copy(sc._style)

        shift          = extra
        row_sum_head   = self.INV_ROW_SUMHEAD  + shift
        row_subtotal   = self.INV_ROW_SUBTOTAL + shift
        row_tax        = self.INV_ROW_TAX      + shift
        row_other      = self.INV_ROW_OTHER    + shift
        row_total      = self.INV_ROW_TOTAL    + shift

        # ---- ヘッダー ----
        ws.cell(self.INV_ROW_NO,       2).value = f'請求書 No.{customer_no}'
        ws.cell(self.INV_ROW_CUSTOMER, 2).value = customer_name
        ws.cell(self.INV_ROW_DATE,     2).value = f'請求日：{invoice_str}'
        ws.cell(self.INV_ROW_DUE,      2).value = f'支払い期日：{due_str}'
        content = data.get('content') or '動画編集案件の件'
        ws.cell(self.INV_ROW_DUE, 3).value = content

        # ---- 明細 ----
        total = 0
        for i in range(max(n, self.INV_DATA_ROWS)):
            r = self.INV_DATA_START + i
            if i < len(cases):
                case = cases[i]
                # 「詳細」は案件名（＝Notionの指定案件ファイル名）を使う。
                # 未設定の場合のみ 備考 → 案件番号 の順にフォールバック。
                desc = case_label(case)
                amount = int(round(float(case.get('amount') or case.get('price') or 0)))
                total += amount
                ws.cell(r, 2).value = desc
                ws.cell(r, 3).value = amount
            else:
                # 雛形の余り行は書式だけ残して中身を空にする
                ws.cell(r, 2).value = None
                ws.cell(r, 3).value = None

        # ---- 集計（雛形と同じ数式のまま行番号だけ追従）----
        ws.cell(row_subtotal, 3).value = '=IFERROR(SUM(InvoiceDetails[金額]), "")'
        ws.cell(row_total,    3).value = f'=IFERROR(C{row_subtotal}*(1+C{row_tax})+C{row_other}, "")'

        # ---- 金額は小数点なし・3桁カンマ区切りに統一 ----
        # 雛形側の書式も直したが、雛形を差し替えられても崩れないよう
        # 出力時に必ず上書きする（税率だけは % 表示のまま）。
        money_rows = list(range(self.INV_DATA_START,
                                self.INV_DATA_START + max(n, self.INV_DATA_ROWS)))
        money_rows += [row_subtotal, row_other, row_total]
        for r in money_rows:
            ws.cell(r, 3).number_format = self.INV_MONEY_FMT
        ws.cell(row_tax, 3).number_format = '0%'

        # ---- テーブル範囲を実データに合わせて更新 ----
        tbl = ws.tables.get('InvoiceDetails')
        if tbl is not None:
            new_ref = f'B12:C{row_sum_head}'
            tbl.ref = new_ref
            if tbl.autoFilter is not None:
                tbl.autoFilter.ref = new_ref

        # ---- 出力 ----
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        wb.save(tmp.name)
        safe = invoice_str.replace('/', '-')
        filename = f'{customer_no}_請求書_{safe}.xlsx'
        print(f"\n📄 請求書生成（雛形ベース）: {filename}  ({len(cases)}件 合計{total:,}円)")
        self.send_file(tmp.name, filename,
                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        os.unlink(tmp.name)

    # ========== カレンダー（案件表と完全同期）==========
    def handle_get_calendar(self):
        from urllib.parse import urlparse, parse_qs
        qs = parse_qs(urlparse(self.path).query)
        month_str = (qs.get('month') or [None])[0]
        today = datetime.date.today()
        if month_str:
            try:
                year  = int(month_str.split('-')[0])
                month = int(month_str.split('-')[1])
            except Exception:
                year, month = today.year, today.month
        else:
            year, month = today.year, today.month

        start = f"{year}-{month:02d}-01"
        end   = f"{year+1}-01-01" if month == 12 else f"{year}-{month+1:02d}-01"
        print(f"\n📅 カレンダー取得: {year}年{month}月")

        events = []
        # 案件表を全件取得（お客様あり=案件、なし=予定）
        q = {"filter": {"and": [
                {"property": "案件締切日・進行", "date": {"on_or_after": start}},
                {"property": "案件締切日・進行", "date": {"before": end}},
             ]}, "page_size": 100}
        cursor = None
        while True:
            if cursor:
                q["start_cursor"] = cursor
            result, _ = notion_request("POST", f"/databases/{CASE_DB_ID}/query", q)
            if not result:
                break
            for page in result.get("results", []):
                try:
                    props    = page["properties"]
                    number   = (props["当方案件番号"]["title"] or [{}])[0].get("plain_text","").strip()
                    customer = (props["お客様no/名"]["rich_text"] or [{}])[0].get("plain_text","").strip()
                    dl       = (props["案件締切日・進行"]["date"] or {}).get("start","")
                    status   = (props["進捗"]["status"] or {}).get("name","")
                    memo_rt  = (props.get("備考") or {}).get("rich_text") or []
                    memo     = memo_rt[0].get("plain_text","") if memo_rt else ""
                    if not dl:
                        continue
                    date_only = dl[:10]
                    # 時刻情報を抽出（例: "2026-03-30T07:30:00.000+09:00" → "07:30"）
                    start_time = dl[11:16] if len(dl) > 10 else ""
                    dl_end = (props["案件締切日・進行"]["date"] or {}).get("end","")
                    end_time = dl_end[11:16] if dl_end and len(dl_end) > 10 else ""
                    if customer:
                        # お客様ありは案件締切
                        events.append({"date": date_only, "title": number,
                                       "type": "deadline", "customer": customer,
                                       "status": status, "memo": memo, "id": page["id"]})
                    else:
                        # お客様なしは予定エントリー（バイト・会議等）
                        events.append({"date": date_only, "title": number,
                                       "type": "schedule", "memo": memo, "id": page["id"],
                                       "startTime": start_time, "endTime": end_time})
                except Exception:
                    pass
            if not result.get("has_more"):
                break
            cursor = result.get("next_cursor")

        self.send_json(200, {"ok": True,
                             "month": f"{year}-{month:02d}",
                             "events": events})

    def handle_get_calendar_day(self):
        """特定日の予定・案件をすべて返す（日次スケジュール用）"""
        from urllib.parse import urlparse, parse_qs
        qs   = parse_qs(urlparse(self.path).query)
        date = (qs.get("date") or [None])[0]
        if not date:
            self.send_json(400, {"ok": False, "error": "dateが必要"}); return
        print(f"\n📅 カレンダー日別取得: {date}")
        next_day = str(datetime.date.fromisoformat(date) + datetime.timedelta(days=1))
        q = {"filter": {"and": [
                {"property": "案件締切日・進行", "date": {"on_or_after": date}},
                {"property": "案件締切日・進行", "date": {"before": next_day}},
             ]}, "page_size": 100}
        result, _ = notion_request("POST", f"/databases/{CASE_DB_ID}/query", q)
        events = []
        for page in (result or {}).get("results", []):
            try:
                props    = page["properties"]
                number   = (props["当方案件番号"]["title"] or [{}])[0].get("plain_text","").strip()
                customer = (props["お客様no/名"]["rich_text"] or [{}])[0].get("plain_text","").strip()
                dl       = (props["案件締切日・進行"]["date"] or {}).get("start","")
                dl_end   = (props["案件締切日・進行"]["date"] or {}).get("end","")
                status   = (props["進捗"]["status"] or {}).get("name","")
                memo_rt  = (props.get("備考") or {}).get("rich_text") or []
                memo     = memo_rt[0].get("plain_text","") if memo_rt else ""
                if not dl: continue
                start_time = dl[11:16] if len(dl) > 10 else ""
                end_time   = dl_end[11:16] if dl_end and len(dl_end) > 10 else ""
                if customer:
                    events.append({"date": dl[:10], "title": number,
                                   "type": "deadline", "customer": customer,
                                   "status": status, "memo": memo, "id": page["id"]})
                else:
                    events.append({"date": dl[:10], "title": number,
                                   "type": "schedule", "memo": memo, "id": page["id"],
                                   "startTime": start_time, "endTime": end_time})
            except Exception:
                pass
        self.send_json(200, {"ok": True, "date": date, "events": events})

    # ========== 引き継ぎ CRUD（Notion DB）==========
    def handle_get_handover_all(self):
        """全アクティブ引き継ぎを日付順で返す"""
        body = {"filter": {"property": "ステータス", "select": {"equals": "active"}},
                "page_size": 100}
        result, err = notion_request("POST", f"/databases/{HANDOVER_DB_ID}/query", body)
        if not result:
            self.send_json(500, {"ok": False, "error": err}); return
        items = [i for i in (parse_handover_page(p) for p in result.get("results",[])) if i]
        items.sort(key=lambda i: i.get("start_date", i.get("date","9999-12-31")))
        self.send_json(200, {"ok": True, "items": items})

    def handle_handover_delete(self, data):
        """引き継ぎをアーカイブ（削除）"""
        item_id = data.get("id","")
        notion_request("PATCH", f"/pages/{item_id}", {"archived": True})
        print(f"  🗑️  引き継ぎ削除: {item_id[:8]}")
        self.send_json(200, {"ok": True})

    def handle_get_handover(self):
        from urllib.parse import urlparse, parse_qs
        qs = parse_qs(urlparse(self.path).query)
        date_str = (qs.get('date') or [None])[0] or datetime.date.today().isoformat()
        items = get_active_handover(date_str)
        self.send_json(200, {"ok": True, "items": items, "date": date_str})

    # ========== メール定型文 ==========
    def handle_mail_templates_list(self):
        """メール定型文を全件取得"""
        items, cursor = [], None
        while True:
            body = {"page_size": 100}
            if cursor:
                body["start_cursor"] = cursor
            result, err = notion_request("POST", f"/databases/{MAIL_TEMPLATE_DB_ID}/query", body)
            if not result:
                self.send_json(400, {"ok": False, "error": err}); return
            for page in result.get("results", []):
                item = parse_mail_template_page(page)
                if item:
                    items.append(item)
            if not result.get("has_more"):
                break
            cursor = result.get("next_cursor")
        # 並び順（未設定は末尾）→ テンプレート名 の順に並べる
        items.sort(key=lambda x: (x["order"] is None,
                                  x["order"] if x["order"] is not None else 0,
                                  x["name"]))
        print(f"\n✉️ メール定型文取得: {len(items)}件")
        self.send_json(200, {"ok": True, "templates": items})

    def handle_mail_template_save(self, data):
        """メール定型文を登録／更新（idがあれば更新）"""
        name = (data.get("name") or "").strip()
        if not name:
            self.send_json(400, {"ok": False, "error": "テンプレート名が必要です"}); return
        props = {
            "テンプレート名": {"title": [{"text": {"content": name[:MAIL_TPL_CHUNK]}}]},
            "件名":          {"rich_text": mail_rt_chunks(data.get("subject", ""))},
            "本文":          {"rich_text": mail_rt_chunks(data.get("body", ""))},
            "メモ":          {"rich_text": mail_rt_chunks(data.get("memo", ""))},
            "有効":          {"checkbox": bool(data.get("enabled", True))},
        }
        cat = (data.get("category") or "").strip()
        props["カテゴリ"] = {"select": {"name": cat}} if cat else {"select": None}
        order = data.get("order")
        if order in (None, ""):
            props["並び順"] = {"number": None}
        else:
            try:
                num = float(order)
                props["並び順"] = {"number": int(num) if num == int(num) else num}
            except (TypeError, ValueError):
                props["並び順"] = {"number": None}

        page_id = data.get("id") or ""
        if page_id:
            result, err = notion_request("PATCH", f"/pages/{page_id}", {"properties": props})
            action = "更新"
        else:
            result, err = notion_request("POST", "/pages",
                                         {"parent": {"database_id": MAIL_TEMPLATE_DB_ID},
                                          "properties": props})
            action = "登録"
        if not result:
            self.send_json(400, {"ok": False, "error": err}); return
        print(f"  ✅ メール定型文{action}: {name[:30]}")
        self.send_json(200, {"ok": True, "template": parse_mail_template_page(result)})

    def handle_mail_template_delete(self, data):
        """メール定型文を削除（Notion上はアーカイブ）"""
        page_id = data.get("id", "")
        if not page_id:
            self.send_json(400, {"ok": False, "error": "idが必要"}); return
        result, err = notion_request("PATCH", f"/pages/{page_id}", {"archived": True})
        if result is None:
            self.send_json(400, {"ok": False, "error": err}); return
        print(f"  🗑 メール定型文削除: {page_id[:8]}")
        self.send_json(200, {"ok": True})

    # ========== 日報 ==========
    def handle_nippo_list(self):
        """日報を全件取得（新しい日付が上）"""
        items, cursor = [], None
        while True:
            body = {"page_size": 100, "sorts": [{"property": "作業日", "direction": "descending"}]}
            if cursor:
                body["start_cursor"] = cursor
            result, err = notion_request("POST", f"/databases/{NIPPO_DB_ID}/query", body)
            if not result:
                self.send_json(400, {"ok": False, "error": err}); return
            for page in result.get("results", []):
                item = parse_nippo_page(page)
                if item:
                    items.append(item)
            if not result.get("has_more"):
                break
            cursor = result.get("next_cursor")
        items.sort(key=lambda x: (x["date"], x["start"]), reverse=True)
        self.send_json(200, {"ok": True, "items": items})

    def handle_nippo_save(self, data):
        """日報を登録／更新（idがあれば更新）"""
        date  = (data.get("date") or "").strip()
        tasks = []
        for t in data.get("tasks") or []:
            if not isinstance(t, dict):
                continue
            memo = str(t.get("memo") or "").rstrip()
            t = {k: str(t.get(k) or "").strip() for k in ("customer", "customerName", "case", "content")}
            t["memo"] = memo
            if any(v.strip() for v in t.values()):
                tasks.append(t)
        title = "／".join(x for x in (nippo_task_label(t) for t in tasks) if x) or "メモ"
        if not tasks:
            self.send_json(400, {"ok": False, "error": "作業を1つ以上入れてください"}); return
        if not date:
            self.send_json(400, {"ok": False, "error": "作業日を入れてください"}); return
        props = {
            "作業内容": {"title": [{"text": {"content": title[:MAIL_TPL_CHUNK]}}]},
            "詳細":     {"rich_text": mail_rt_chunks("\n\n".join(
                            (f"■{nippo_task_label(t)}\n" if nippo_task_label(t) else "") + t["memo"]
                            for t in tasks if t["memo"].strip()) or data.get("detail", ""))},
            "作業日":   {"date": {"start": date}},
            "開始":     {"rich_text": mail_rt_chunks((data.get("start") or "").strip())},
            "終了":     {"rich_text": mail_rt_chunks((data.get("end") or "").strip())},
            "場所":     {"rich_text": mail_rt_chunks((data.get("place") or "").strip())},
            "タスク":   {"rich_text": mail_rt_chunks(jsonlib.dumps(tasks, ensure_ascii=False))},
        }
        page_id = data.get("id") or ""
        if page_id:
            result, err = notion_request("PATCH", f"/pages/{page_id}", {"properties": props})
        else:
            result, err = notion_request("POST", "/pages",
                                         {"parent": {"database_id": NIPPO_DB_ID}, "properties": props})
        if not result:
            self.send_json(400, {"ok": False, "error": err}); return
        print(f"  📝 日報{'更新' if page_id else '登録'}: {date} {title[:30]}")
        self.send_json(200, {"ok": True, "item": parse_nippo_page(result)})

    def handle_nippo_delete(self, data):
        """日報を削除（Notion上はアーカイブ）"""
        page_id = data.get("id", "")
        if not page_id:
            self.send_json(400, {"ok": False, "error": "idが必要"}); return
        result, err = notion_request("PATCH", f"/pages/{page_id}", {"archived": True})
        if result is None:
            self.send_json(400, {"ok": False, "error": err}); return
        self.send_json(200, {"ok": True})

    # ========== 営業ツール ==========
    def handle_eigyo(self, path, data):
        """営業ツールの API（すべて管理者ログインが必要）"""
        if not ADMIN_TOKEN or data.get("token", "") != ADMIN_TOKEN:
            self.send_json(401, {"ok": False, "error": "ポータルの管理者ログインが必要です"}); return
        act = path[len("/api/eigyo/"):]
        if act == "list":
            pages, err = notion_query_all(EIGYO_LIST_DB_ID, {"sorts": [{"timestamp": "created_time", "direction": "ascending"}]})
            if pages is None:
                self.send_json(400, {"ok": False, "error": err}); return
            items = [x for x in map(parse_eigyo_page, pages) if x]
            self.send_json(200, {"ok": True, "items": items, "dupGroups": eigyo_dup_groups(items)})
        elif act == "save":
            x = data.get("item") or {}
            if not (x.get("company") or x.get("site") or x.get("email")):
                self.send_json(400, {"ok": False, "error": "会社名・サイト・メールのどれかは入れてください"}); return
            if not x.get("id") and not data.get("force"):
                keys = set(eigyo_key(x))
                pages, _ = notion_query_all(EIGYO_LIST_DB_ID)
                dups = [y for y in map(parse_eigyo_page, pages or []) if y and keys & set(eigyo_key(y))]
                if dups:
                    self.send_json(200, {"ok": False, "duplicate": True, "matches": dups}); return
            if x.get("id"):
                result, err = notion_request("PATCH", f"/pages/{x['id']}", {"properties": eigyo_props(x)})
            else:
                result, err = notion_request("POST", "/pages", {"parent": {"database_id": EIGYO_LIST_DB_ID}, "properties": eigyo_props(x)})
            if not result:
                self.send_json(400, {"ok": False, "error": err}); return
            self.send_json(200, {"ok": True, "item": parse_eigyo_page(result)})
        elif act == "merge":
            # 重複をまとめる。ids を渡せばそのグループだけ、無ければ全グループ
            pages, err = notion_query_all(EIGYO_LIST_DB_ID)
            if pages is None:
                self.send_json(400, {"ok": False, "error": err}); return
            items = {x["id"]: x for x in map(parse_eigyo_page, pages) if x}
            groups = [data["ids"]] if data.get("ids") else eigyo_dup_groups(list(items.values()))
            merged = removed = 0
            for g in groups:
                rows = [items[i] for i in g if i in items]
                if len(rows) < 2:
                    continue
                base = eigyo_merge(rows)
                notion_request("PATCH", f"/pages/{base['id']}", {"properties": eigyo_props(base)})
                for r in rows:
                    if r["id"] != base["id"]:
                        notion_request("PATCH", f"/pages/{r['id']}", {"archived": True}); removed += 1
                merged += 1
            self.send_json(200, {"ok": True, "merged": merged, "removed": removed})
        elif act == "delete":
            ids = data.get("ids") or ([data["id"]] if data.get("id") else [])
            for pid in ids:
                notion_request("PATCH", f"/pages/{pid}", {"archived": True})
            self.send_json(200, {"ok": True, "count": len(ids)})
        elif act == "status":
            ids, st = data.get("ids") or [], data.get("status", "")
            if st not in EIGYO_STATUSES:
                self.send_json(400, {"ok": False, "error": "ステータスが不正です"}); return
            for pid in ids:
                notion_request("PATCH", f"/pages/{pid}", {"properties": {"ステータス": {"select": {"name": st}}}})
            self.send_json(200, {"ok": True, "count": len(ids)})
        elif act == "import":
            # 画面側で少しずつ分けて送ってくる。dupMode: skip=既存は飛ばす / fill=既存の空欄だけ埋める / add=別の行として入れる
            rows = data.get("rows") or []
            mode = data.get("dupMode") or ("skip" if data.get("skipDup", True) else "add")
            owner = {}
            if mode != "add":
                pages, err = notion_query_all(EIGYO_LIST_DB_ID)
                if pages is None:
                    self.send_json(400, {"ok": False, "error": err}); return
                for x in map(parse_eigyo_page, pages):
                    if x:
                        for k in eigyo_key(x):
                            owner.setdefault(k, x)
            added, skipped, filled, failed = 0, 0, 0, []
            for r in rows:
                r["line"], r["memo"] = eigyo_split_line(r.get("memo"), r.get("line"))
                keys = eigyo_key(r)
                hit = next((owner[k] for k in keys if k in owner), None) if mode != "add" else None
                if hit:
                    if mode == "fill":
                        patch = {}
                        for f in ("industry", "site", "form", "email", "person", "line", "memo"):
                            if not (hit.get(f) or "").strip() and (r.get(f) or "").strip():
                                hit[f] = r[f]; patch[f] = True
                        if EIGYO_STATUS_RANK.get(r.get("status"), 0) > EIGYO_STATUS_RANK.get(hit.get("status"), 0):
                            hit["status"] = r["status"]; patch["status"] = True
                        if patch:
                            props = {k: v for k, v in eigyo_props(hit).items()
                                     if k in ("業種", "サイトURL", "フォームURL", "メール", "担当者", "個別の一文", "メモ", "ステータス")}
                            result, err = notion_request("PATCH", f"/pages/{hit['id']}", {"properties": props})
                            if result:
                                filled += 1
                            else:
                                failed.append(f"{r.get('company', '')}: {err}")
                            _time.sleep(0.34)
                        else:
                            skipped += 1
                    else:
                        skipped += 1
                    continue
                result, err = notion_request("POST", "/pages", {"parent": {"database_id": EIGYO_LIST_DB_ID}, "properties": eigyo_props(r)})
                if result:
                    added += 1
                    x = parse_eigyo_page(result)
                    if x and mode != "add":
                        for k in keys:
                            owner.setdefault(k, x)
                else:
                    failed.append(f"{r.get('company', '')}: {err}")
                _time.sleep(0.34)          # Notion の流量制限（3件/秒）
            self.send_json(200, {"ok": True, "added": added, "skipped": skipped, "filled": filled, "failed": failed})
        elif act == "templates":
            pages, err = notion_query_all(EIGYO_TPL_DB_ID)
            if pages is None:
                self.send_json(400, {"ok": False, "error": err}); return
            items = [x for x in map(parse_eigyo_tpl_page, pages) if x]
            items.sort(key=lambda x: (x["order"] is None, x["order"] or 0, x["name"]))
            self.send_json(200, {"ok": True, "templates": items})
        elif act == "templates/save":
            t = data.get("template") or {}
            if not (t.get("name") or "").strip():
                self.send_json(400, {"ok": False, "error": "テンプレート名を入れてください"}); return
            props = {
                "テンプレート名": {"title": [{"text": {"content": t["name"].strip()[:MAIL_TPL_CHUNK]}}]},
                "業種": {"rich_text": mail_rt_chunks((t.get("industry") or "").strip())},
                "件名": {"rich_text": mail_rt_chunks(t.get("subject") or "")},
                "本文": {"rich_text": mail_rt_chunks(t.get("body") or "")},
                "有効": {"checkbox": bool(t.get("enabled", True))},
                "並び順": {"number": t.get("order") if isinstance(t.get("order"), (int, float)) else None},
            }
            if t.get("id"):
                result, err = notion_request("PATCH", f"/pages/{t['id']}", {"properties": props})
            else:
                result, err = notion_request("POST", "/pages", {"parent": {"database_id": EIGYO_TPL_DB_ID}, "properties": props})
            if not result:
                self.send_json(400, {"ok": False, "error": err}); return
            self.send_json(200, {"ok": True, "template": parse_eigyo_tpl_page(result)})
        elif act == "templates/delete":
            notion_request("PATCH", f"/pages/{data.get('id', '')}", {"archived": True})
            self.send_json(200, {"ok": True})
        elif act == "mail-status":
            self.send_json(200, {"ok": True, "configured": bool(GMAIL_USER and GMAIL_APP_PASSWORD),
                                 "user": GMAIL_USER, "sender": EIGYO_SENDER_NAME,
                                 "today": self._eigyo_mail_count_today(), "limit": EIGYO_DAILY_LIMIT})
        elif act == "send-mail":
            self.handle_eigyo_send_mail(data)
        elif act == "auto":
            op = data.get("op") or "status"
            if op == "on":
                if not (GMAIL_USER and GMAIL_APP_PASSWORD) and not EIGYO_AUTO_DRY:
                    self.send_json(400, {"ok": False, "error": "Railway に GMAIL_USER と GMAIL_APP_PASSWORD を設定してください"}); return
                EIGYO_AUTO["enabled"] = True; _eigyo_auto_save(); _eigyo_auto_log("設定", "自動送信をONにしました")
            elif op == "off":
                EIGYO_AUTO["enabled"] = False; _eigyo_auto_save(); _eigyo_auto_log("設定", "自動送信をOFFにしました")
            elif op == "run-one":
                ok, msg = eigyo_auto_send_one("手動")
                self.send_json(200, {"ok": ok, "message": msg, "status": eigyo_auto_status(),
                                     "preview": EIGYO_AUTO.get("last_mail") if ok and EIGYO_AUTO_DRY else None}); return
            elif op == "refresh":
                pages, _ = notion_query_all(EIGYO_LIST_DB_ID)
                EIGYO_AUTO["remaining"] = len(eigyo_auto_candidates([y for y in map(parse_eigyo_page, pages or []) if y]))
            self.send_json(200, {"ok": True, "status": eigyo_auto_status()})
        else:
            self.send_json(404, {"ok": False, "error": "not found"})

    def _eigyo_mail_count_today(self):
        return _eigyo_count_today()

    def handle_eigyo_send_mail(self, data):
        """営業メールを1通送り、リストを送信済にする（画面の「送る」ボタン用）"""
        ok, res = eigyo_send_core(data.get("to"), data.get("subject"), data.get("body"), pid=data.get("id") or "",
                                  template_name=data.get("template") or "", force=bool(data.get("force")))
        if not ok:
            msg = str(res).replace("AUTH:", "")
            self.send_json(400, {"ok": False, "duplicate": "送信済みです" in msg, "error": msg}); return
        self.send_json(200, {"ok": True, "today": self._eigyo_mail_count_today(), "limit": EIGYO_DAILY_LIMIT})

    def handle_handover_add(self, data):
        text = data.get("text","").strip()
        typ  = data.get("type","todo")
        if not text:
            self.send_json(400, {"ok": False, "error": "テキストが必要"}); return
        today    = datetime.date.today().isoformat()
        tomorrow = (datetime.date.today() + datetime.timedelta(days=1)).isoformat()
        props = {
            "内容":     {"title":  [{"text": {"content": text}}]},
            "種別":     {"select": {"name": typ}},
            "ステータス": {"select": {"name": "active"}},
            "作成日":   {"date":   {"start": today}},
        }
        if typ == "medium":
            props["開始日"] = {"date": {"start": data.get("start_date", today)}}
            props["終了日"] = {"date": {"start": data.get("end_date",   today)}}
        else:
            props["日付"] = {"date": {"start": data.get("date", tomorrow)}}
        result, err = notion_request("POST", "/pages",
                                     {"parent": {"database_id": HANDOVER_DB_ID},
                                      "properties": props})
        if not result:
            self.send_json(400, {"ok": False, "error": err}); return
        item = parse_handover_page(result)
        print(f"  ✅ 引き継ぎ追加: [{typ}] {text[:30]}")
        self.send_json(200, {"ok": True, "item": item})

    def handle_handover_done(self, data):
        item_id = data.get("id","")
        notion_request("PATCH", f"/pages/{item_id}",
                       {"properties": {"ステータス": {"select": {"name": "done"}}}})
        print(f"  ✅ 引き継ぎ完了: {item_id[:8]}")
        self.send_json(200, {"ok": True})

    def handle_handover_carry(self, data):
        """翌日へ引き継ぎ（種別を変えて再登録）"""
        item_id = data.get("id","")
        as_type = data.get("as_type","todo")
        to_date = data.get("to_date",
                           (datetime.date.today() + datetime.timedelta(days=1)).isoformat())
        # 元アイテムを取得して完了にする
        src_result, _ = notion_request("GET", f"/pages/{item_id}")
        notion_request("PATCH", f"/pages/{item_id}",
                       {"properties": {"ステータス": {"select": {"name": "done"}}}})
        if src_result:
            src = parse_handover_page(src_result)
            if src:
                props = {
                    "内容":     {"title":  [{"text": {"content": src["text"]}}]},
                    "種別":     {"select": {"name": as_type}},
                    "ステータス": {"select": {"name": "active"}},
                    "作成日":   {"date":   {"start": datetime.date.today().isoformat()}},
                    "日付":     {"date":   {"start": to_date}},
                }
                notion_request("POST", "/pages",
                               {"parent": {"database_id": HANDOVER_DB_ID},
                                "properties": props})
                print(f"  🔁 引き継ぎ → {to_date} [{as_type}] {src['text'][:30]}")
        self.send_json(200, {"ok": True})

    def handle_handover_extend(self, data):
        """中期メモの終了日を延長"""
        item_id = data.get("id","")
        new_end  = data.get("end_date","")
        notion_request("PATCH", f"/pages/{item_id}",
                       {"properties": {"終了日": {"date": {"start": new_end}}}})
        print(f"  📅 中期延長: {item_id[:8]} → {new_end}")
        self.send_json(200, {"ok": True})

    def handle_handover_update_date(self, data):
        """日付変更"""
        item_id = data.get("id","")
        props = {}
        if data.get("start_date"): props["開始日"] = {"date": {"start": data["start_date"]}}
        if data.get("end_date"):   props["終了日"] = {"date": {"start": data["end_date"]}}
        if data.get("date"):       props["日付"]   = {"date": {"start": data["date"]}}
        if props:
            notion_request("PATCH", f"/pages/{item_id}", {"properties": props})
        print(f"  📅 日付変更: {item_id[:8]}")
        self.send_json(200, {"ok": True})

    def handle_get_handover_done_list(self):
        """完了済み引き継ぎ一覧を返す"""
        items = get_done_handover()
        self.send_json(200, {"ok": True, "items": items})

    def handle_handover_restore(self, data):
        """完了済み引き継ぎをアクティブに戻す"""
        item_id = data.get("id", "")
        if not item_id:
            self.send_json(400, {"ok": False, "error": "idが必要"}); return
        notion_request("PATCH", f"/pages/{item_id}",
                       {"properties": {"ステータス": {"select": {"name": "active"}}}})
        print(f"  🔄 引き継ぎ復元: {item_id[:8]}")
        self.send_json(200, {"ok": True})

    def handle_handover_update_content(self, data):
        """引き継ぎ内容テキストを更新"""
        item_id = data.get("id", "")
        text    = data.get("text", "").strip()
        if not item_id or not text:
            self.send_json(400, {"ok": False, "error": "idとtextが必要"}); return
        notion_request("PATCH", f"/pages/{item_id}", {
            "properties": {"内容": {"title": [{"text": {"content": text}}]}}
        })
        print(f"  ✏️  内容更新: {item_id[:8]} → {text[:20]}")
        self.send_json(200, {"ok": True})

    def handle_get_schedule_dates(self):
        """指定月に日次スケジュールが保存されている日付一覧を返す"""
        from urllib.parse import urlparse, parse_qs
        qs    = parse_qs(urlparse(self.path).query)
        month = (qs.get("month") or [None])[0]  # "YYYY-MM"
        if not month:
            self.send_json(400, {"ok": False, "error": "monthが必要"}); return
        schedules = load_daily_schedules()
        # ブロックが1件以上ある日のみ返す
        dates = [d for d in schedules if d.startswith(month) and len(schedules[d]) > 0]
        self.send_json(200, {"ok": True, "dates": dates})

    def handle_get_daily_schedule(self):
        """日次スケジュールブロックを返す"""
        from urllib.parse import urlparse, parse_qs
        qs   = parse_qs(urlparse(self.path).query)
        date = qs.get("date", [None])[0]
        if not date:
            self.send_json(400, {"ok": False, "error": "dateが必要"}); return
        schedules = load_daily_schedules()
        blocks = schedules.get(date, [])
        self.send_json(200, {"ok": True, "date": date, "blocks": blocks})

    def handle_save_daily_block(self, data):
        """日次スケジュールブロックを保存/更新"""
        import uuid
        date  = data.get("date", "")
        block = data.get("block", {})
        if not date or not block:
            self.send_json(400, {"ok": False, "error": "dateとblockが必要"}); return
        if not block.get("id"):
            block["id"] = str(uuid.uuid4())[:8]
        schedules = load_daily_schedules()
        blocks = schedules.get(date, [])
        # 同IDがあれば更新、なければ追加
        updated = False
        for i, b in enumerate(blocks):
            if b.get("id") == block["id"]:
                blocks[i] = block; updated = True; break
        if not updated:
            blocks.append(block)
        schedules[date] = blocks
        save_daily_schedules(schedules)
        print(f"  📅 日次ブロック保存: {date} id={block['id']}")
        self.send_json(200, {"ok": True, "block": block})

    def handle_delete_daily_block(self, data):
        """日次スケジュールブロックを削除"""
        date     = data.get("date", "")
        block_id = data.get("id", "")
        if not date or not block_id:
            self.send_json(400, {"ok": False, "error": "dateとidが必要"}); return
        schedules = load_daily_schedules()
        blocks = schedules.get(date, [])
        schedules[date] = [b for b in blocks if b.get("id") != block_id]
        save_daily_schedules(schedules)
        print(f"  🗑️  日次ブロック削除: {date} id={block_id}")
        self.send_json(200, {"ok": True})

    def handle_move_notion_event(self, data):
        """Notion予定ブロックの時刻を更新"""
        page_id   = data.get("id", "")
        date      = data.get("date", "")
        start_min = int(data.get("startMin", 0))
        end_min   = int(data.get("endMin", 60))
        if not page_id or not date:
            self.send_json(400, {"ok": False, "error": "id/dateが必要"}); return
        sh = f"{start_min // 60:02d}:{start_min % 60:02d}"
        eh = f"{end_min   // 60:02d}:{end_min   % 60:02d}"
        start_dt = f"{date}T{sh}:00+09:00"
        end_dt   = f"{date}T{eh}:00+09:00"
        props = {"案件締切日・進行": {"date": {"start": start_dt, "end": end_dt}}}
        result, err = notion_request("PATCH", f"/pages/{page_id}", {"properties": props})
        if err:
            self.send_json(500, {"ok": False, "error": err}); return
        print(f"  📅 Notion予定移動: {page_id[:8]} {sh}〜{eh}")
        self.send_json(200, {"ok": True, "startTime": sh, "endTime": eh})

    def handle_get_cases_list(self):
        """全案件を検索・フィルタして返す（案件履歴ツール用）"""
        from urllib.parse import urlparse, parse_qs
        qs = parse_qs(urlparse(self.path).query)
        q_term   = (qs.get("q")        or [""])[0].strip()
        status_f = (qs.get("status")   or [""])[0].strip()
        cust_f   = (qs.get("customer") or [""])[0].strip()

        month_f  = (qs.get("month")    or [""])[0].strip()   # e.g. "2026-03"

        filters = []
        # 月フィルター（デフォルト: 当月）
        if month_f:
            try:
                y, m = map(int, month_f.split('-'))
                ms = f"{y}-{m:02d}-01"
                me = f"{y+1}-01-01" if m == 12 else f"{y}-{m+1:02d}-01"
                filters.append({"property": "案件締切日・進行", "date": {"on_or_after": ms}})
                filters.append({"property": "案件締切日・進行", "date": {"before": me}})
            except Exception:
                pass
        if cust_f:
            filters.append({"property": "お客様no/名", "rich_text": {"contains": cust_f}})
        if status_f:
            filters.append({"property": "進捗", "status": {"equals": status_f}})
        # 最終確認の済／未で絞り込む
        conf_f = (qs.get("confirmed") or [""])[0].strip()
        if conf_f in ("1", "0"):
            filters.append({"property": "最終確認完了", "checkbox": {"equals": conf_f == "1"}})
        # キーワード検索は案件番号のみ（Notion APIのorフィルター制限を回避）
        if q_term:
            filters.append({"property": "当方案件番号", "title": {"contains": q_term}})

        query_body = {
            "page_size": 100,
            "sorts": [{"property": "案件締切日・進行", "direction": "descending"}],
        }
        if len(filters) == 1:   query_body["filter"] = filters[0]
        elif len(filters) > 1:  query_body["filter"] = {"and": filters}

        cases = []
        cursor = None
        while True:
            if cursor: query_body["start_cursor"] = cursor
            result, err = notion_request("POST", f"/databases/{CASE_DB_ID}/query", query_body)
            if not result:
                self.send_json(400, {"ok": False, "error": err}); return
            for page in result.get("results", []):
                try:
                    props    = page["properties"]
                    number   = (props["当方案件番号"]["title"]      or [{}])[0].get("plain_text","").strip()
                    customer = (props["お客様no/名"]["rich_text"]   or [{}])[0].get("plain_text","").strip()
                    price_v  = props["単価"]["number"]
                    price    = price_v if price_v is not None else 0
                    # 案件名は「指定案件ファイル名」。備考は「備考」1本に統合（2026-09-01）
                    filename = (props.get("指定案件ファイル名",{}).get("rich_text") or [{}])[0].get("plain_text","")
                    memo     = (props.get("備考",{}).get("rich_text")        or [{}])[0].get("plain_text","")
                    note     = memo
                    gross    = (props.get("粗利（単価-外注費）",{}).get("rich_text") or [{}])[0].get("plain_text","")
                    dl       = (props.get("案件締切日・進行",{}).get("date") or {}).get("start","")
                    status   = (props.get("進捗",{}).get("status") or {}).get("name","")
                    # 請求後の最終確認が済んだか。進捗ステータスとは別管理。
                    # 進捗に「請求済み」を足すとNotionの選択肢にないためエラーになるので、
                    # チェックボックスの独立項目にしている（2026-09-01）
                    confirmed = bool(props.get("最終確認完了",{}).get("checkbox") or False)
                    # 粗利テキストから利益額を抽出（例: "10,000 - 3,000 = 7,000" → 7000）
                    # grossが空の場合は外注費ゼロとみなし粗利=単価として扱う
                    import re as _re
                    profit = None
                    if gross:
                        m = _re.search(r'=\s*([\d,]+)', gross)
                        if m:
                            try: profit = int(m.group(1).replace(',',''))
                            except: pass
                    if profit is None and price:
                        profit = int(price)   # 外注費未入力 → 粗利=単価
                    if not customer:
                        continue   # 客番号なし（バイト・会議等の予定）は除外
                    cases.append({
                        "id": page["id"], "number": number, "customer": customer,
                        "price": price, "note": note, "filename": filename,
                        "memo": memo, "gross": gross, "profit": profit,
                        "confirmed": confirmed,
                        "date": dl[:10] if dl else "", "status": status,
                        "url": page.get("url",""),
                    })
                except Exception:
                    pass
            if not result.get("has_more"): break
            cursor = result.get("next_cursor")

        # 顧客名を付与（キャッシュから引く）
        name_map = get_customer_name_map()
        for c in cases:
            c["customerName"] = name_map.get(c["customer"], "")

        print(f"\n📋 案件履歴取得: {len(cases)}件 (q={q_term!r} status={status_f!r} cust={cust_f!r})")
        self.send_json(200, {"ok": True, "cases": cases, "total": len(cases)})

    def handle_update_case(self, data):
        """案件のプロパティを更新"""
        page_id = data.get("id","")
        if not page_id:
            self.send_json(400, {"ok": False, "error": "idが必要"}); return
        props = {}
        if "number"   in data: props["当方案件番号"]        = {"title": [{"text": {"content": data["number"]}}]}
        if "customer" in data: props["お客様no/名"]         = {"rich_text": [{"text": {"content": data["customer"]}}]}
        # 備考は1項目に統合済み。note / memo どちらの名前で来ても同じ「備考」へ。
        # 同じプロパティを2行に分けて書くと後勝ちで片方が消えるため1箇所にまとめる。
        if "note" in data or "memo" in data:
            memo_v = data.get("memo") if "memo" in data else data.get("note")
            props["備考"] = {"rich_text": [{"text": {"content": memo_v or ""}}]}
        if "filename" in data: props["指定案件ファイル名"]  = {"rich_text": [{"text": {"content": data["filename"]}}]}
        # 最終確認完了。クライアントから戻しがあればチェックを外して差し戻せる。
        if "confirmed" in data: props["最終確認完了"] = {"checkbox": bool(data["confirmed"])}
        if "status"   in data: props["進捗"]                = {"status": {"name": data["status"]}}
        if "date"     in data:
            props["案件締切日・進行"] = {"date": {"start": data["date"]}} if data["date"] else {"date": None}
        if "price" in data:
            try:
                pv = float(data["price"]) if str(data["price"]).strip() != "" else 0.0
                props["単価"] = {"number": pv}
                # 外注費：空・未送信はすべて0扱い
                raw_cost = str(data.get("outsourceCost", "")).strip()
                cv = float(raw_cost) if raw_cost != "" else 0.0
                gross_text = f"{int(pv):,} - {int(cv):,} = {int(pv - cv):,}"
                props["粗利（単価-外注費）"] = {"rich_text": [{"text": {"content": gross_text}}]}
            except Exception as e:
                print(f"  ⚠️ 粗利計算エラー: {e}")
        if not props:
            self.send_json(400, {"ok": False, "error": "更新項目なし"}); return
        result, err = notion_request("PATCH", f"/pages/{page_id}", {"properties": props})
        if result:
            print(f"  ✅ 案件更新: {page_id[:8]} → {list(props.keys())}")
            self.send_json(200, {"ok": True})
        else:
            self.send_json(500, {"ok": False, "error": err})

    def handle_gantt_excel(self, data):
        """ガントチャートExcelを生成して返す"""
        import tempfile
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        date_str = data.get("date", "")
        blocks   = data.get("blocks", [])
        notion   = data.get("notionEvents", [])

        def time_to_min(t):
            try:
                h, m = t.split(":")
                return int(h) * 60 + int(m)
            except Exception:
                return 0

        SLOT_MINS = 30
        SLOTS     = 24 * 60 // SLOT_MINS  # 48列

        TYPE_COLORS = {
            "custom":   "7C3AED",
            "todo":     "2563EB",
            "memo":     "D97706",
            "deadline": "DC2626",
            "medium":   "059669",
            "notion":   "374151",
        }
        TYPE_LABELS = {
            "custom": "作業", "todo": "TODO", "memo": "メモ",
            "deadline": "納期", "medium": "中期", "notion": "Notion予定",
        }

        entries = []
        for b in blocks:
            entries.append({
                "text":  b.get("text") or "(無題)",
                "type":  b.get("type", "custom"),
                "start": b.get("startMin", 0),
                "end":   b.get("endMin",   b.get("startMin", 0) + 60),
                "notion_linked": bool(b.get("notionId")),
            })
        for ev in notion:
            st = time_to_min(ev.get("startTime", ""))
            en = time_to_min(ev.get("endTime",   "")) if ev.get("endTime") else st + 60
            entries.append({
                "text":  ev.get("title") or "(無題)",
                "type":  "notion",
                "start": st,
                "end":   en,
                "notion_linked": True,
            })
        entries.sort(key=lambda e: e["start"])

        wb = Workbook()
        ws = wb.active
        ws.title = "ガントチャート"

        HEADER_FILL = PatternFill("solid", start_color="1E1B4B", end_color="1E1B4B")
        HOUR_FILL   = PatternFill("solid", start_color="312E81", end_color="312E81")
        EVEN_ROW    = PatternFill("solid", start_color="F8F9FF", end_color="F8F9FF")
        ODD_ROW     = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
        thin        = Side(border_style="thin", color="D1D5DB")
        bdr         = Border(left=thin, right=thin, top=thin, bottom=thin)

        def af(sz=9, bold=False, color="FFFFFF"):
            return Font(name="Arial", size=sz, bold=bold, color=color)

        # 列オフセット：A=内容, B=種別, C=時刻, D以降=ガントバー
        BAR_OFFSET = 4   # ガントバーが始まる列番号（1-indexed）

        # 1行目：タイトル（A〜最終ガント列まで結合）
        ws.merge_cells(f"A1:{get_column_letter(SLOTS + BAR_OFFSET - 1)}1")
        tc = ws["A1"]
        tc.value     = f"日次スケジュール　ガントチャート　{date_str}"
        tc.font      = af(11, True)
        tc.fill      = HEADER_FILL
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        # 2行目：固定ヘッダー（内容／種別／時刻）
        for col, label in [(1, "内容"), (2, "種別"), (3, "時刻")]:
            c = ws.cell(2, col)
            c.value     = label
            c.font      = af(9, True)
            c.fill      = HEADER_FILL
            c.alignment = Alignment(horizontal="center", vertical="center")

        # 2行目：24時間軸ラベル（1時間ごとに表示）
        for i in range(SLOTS):
            mins = i * SLOT_MINS
            h, m = divmod(mins, 60)
            c = ws.cell(2, i + BAR_OFFSET)
            c.value     = f"{h:02d}:00" if m == 0 else ""
            c.font      = af(8, bold=(m == 0))
            c.fill      = HOUR_FILL if m == 0 else HEADER_FILL
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 16

        # 列幅
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 9
        ws.column_dimensions["C"].width = 13   # 時刻列
        for i in range(SLOTS):
            ws.column_dimensions[get_column_letter(i + BAR_OFFSET)].width = 2.2

        # データ行
        for ri, entry in enumerate(entries):
            row   = ri + 3
            bfill = EVEN_ROW if ri % 2 == 0 else ODD_ROW
            color = TYPE_COLORS.get(entry["type"], "7C3AED")
            gfill = PatternFill("solid", start_color=color, end_color=color)

            # 内容
            nc = ws.cell(row, 1)
            nc.value     = ("📅 " if entry["notion_linked"] else "") + entry["text"]
            nc.font      = Font(name="Arial", size=9)
            nc.fill      = bfill
            nc.alignment = Alignment(vertical="center")
            nc.border    = bdr

            # 種別
            sc = ws.cell(row, 2)
            sc.value     = TYPE_LABELS.get(entry["type"], entry["type"])
            sc.font      = Font(name="Arial", size=8)
            sc.fill      = bfill
            sc.alignment = Alignment(horizontal="center", vertical="center")
            sc.border    = bdr

            # 時刻（専用列）
            sh = f"{entry['start'] // 60:02d}:{entry['start'] % 60:02d}"
            eh = f"{entry['end']   // 60:02d}:{entry['end']   % 60:02d}"
            tc = ws.cell(row, 3)
            tc.value     = f"{sh}〜{eh}"
            tc.font      = Font(name="Arial", size=9, bold=True, color="111827")
            tc.fill      = bfill
            tc.alignment = Alignment(horizontal="center", vertical="center")
            tc.border    = bdr

            # ガントバー
            start_slot = entry["start"] // SLOT_MINS
            end_slot   = (entry["end"] + SLOT_MINS - 1) // SLOT_MINS
            for s in range(SLOTS):
                c = ws.cell(row, s + BAR_OFFSET)
                c.fill   = gfill if start_slot <= s < end_slot else bfill
                c.border = bdr
            ws.row_dimensions[row].height = 16

        if not entries:
            ws.cell(3, 1).value = "予定なし"
            ws.cell(3, 1).font  = Font(name="Arial", size=9, color="9CA3AF")

        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        wb.save(tmp.name)

        safe_date = date_str.replace("-", "")
        filename  = f"ガントチャート_{safe_date}.xlsx"
        print(f"\n📊 ガントチャート生成: {filename}  ({len(entries)}件)")
        self.send_file(tmp.name, filename,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        os.unlink(tmp.name)

    def handle_invoice_numbers(self, data):
        """Numbers.appにAppleScriptファイル経由でテンプレートを書き換えて請求書を生成"""
        import shutil, calendar, re, subprocess, tempfile

        if not shutil.which("osascript"):
            self.send_json(500, {"ok": False, "error": "osascriptが見つかりません。macOS上でサーバーを起動してください"}); return

        template_path = os.path.join(SCRIPT_DIR, "請求書雛形のコピー.numbers")
        if not os.path.exists(template_path):
            self.send_json(404, {"ok": False, "error": "請求書雛形のコピー.numbersが見つかりません"}); return

        customer_name = data.get("customerName", "")
        invoice_date  = data.get("invoiceDate", "")
        cases         = data.get("cases", [])
        customer_no   = data.get("customerNo", "")

        # 支払い期日（翌月末）
        due_date = ""
        md = re.match(r"(\d+)年(\d+)月(\d+)日", invoice_date)
        if md:
            y, mo = int(md.group(1)), int(md.group(2))
            nm = mo + 1 if mo < 12 else 1
            ny = y if mo < 12 else y + 1
            due_date = f"{ny}年{nm}月{calendar.monthrange(ny, nm)[1]}日"

        content_parts = [case_label(c) for c in cases if case_label(c)]
        content_str = "、".join(content_parts) or "動画編集案件の件"

        total = sum(int(float(c.get("amount") or c.get("price") or 0)) for c in cases)

        def qs(s):
            """AppleScript文字列用クォートエスケープ"""
            return str(s).replace("\\", "\\\\").replace('"', '\\"')

        # テンプレート構造: データ行13〜14、小計15行目
        # 3件以上の場合は小計行(15)の直前に1行ずつ挿入
        TEMPLATE_ROWS = 2
        extra_rows   = max(0, len(cases) - TEMPLATE_ROWS)
        subtotal_row = 15 + extra_rows

        # 行挿入ブロック（挿入ごとに delay 0.5 で Numbers が処理を完了するのを待つ）
        insert_lines = []
        for i in range(extra_rows):
            insert_lines.append(f'                    add row above row {15 + i}')
            insert_lines.append(f'                    delay 0.5')
        insert_block = "\n".join(insert_lines) + "\n" if insert_lines else ""

        # データ行セット
        row_lines = []
        for i, c in enumerate(cases):
            row    = 13 + i
            detail = qs(case_label(c))
            amount = int(float(c.get("amount") or c.get("price") or 0))
            row_lines.append(f'                    set value of cell "B{row}" to "{detail}"')
            row_lines.append(f'                    set value of cell "C{row}" to {amount}')
        row_block = "\n".join(row_lines)

        # 一時ディレクトリ（ASCIIパス）
        tmp_dir   = tempfile.mkdtemp(prefix="inv_")
        tmp_path  = os.path.join(tmp_dir, "invoice.numbers")
        scpt_path = os.path.join(tmp_dir, "fill.applescript")
        shutil.copy(template_path, tmp_path)

        # AppleScriptをUTF-8ファイルとして書き出し（-e フラグだと日本語が文字化けするため）
        scpt = f'''\
-- Numbers請求書自動記入スクリプト
with timeout of 120 seconds
    set docFile to POSIX file "{tmp_path}"
    tell application "Numbers"
        set wasRunning to running
        set theDoc to open docFile
        -- ファイルが完全に開くまで待機
        delay 5
        tell theDoc
            tell sheet 1
                tell table 1
{insert_block}                    set value of cell "B2" to "{qs(customer_name)}様"
                    set value of cell "B8" to "請求日：{qs(invoice_date)}"
                    set value of cell "B9" to "支払い期日：{qs(due_date)}"
                    set value of cell "C9" to "{qs(content_str)}"
{row_block}
                    set value of cell "C{subtotal_row}" to {total}
                end tell
            end tell
            save
            delay 1
        end tell
        close theDoc saving no
        if not wasRunning then quit
    end tell
end timeout
'''
        with open(scpt_path, "w", encoding="utf-8") as f:
            f.write(scpt)

        print(f"\n📄 Numbers AppleScript:\n{scpt}")

        try:
            result = subprocess.run(
                ["osascript", scpt_path],
                capture_output=True, text=True, timeout=150
            )
            if result.stdout.strip():
                print(f"  stdout: {result.stdout.strip()}")
            if result.stderr.strip():
                print(f"  stderr: {result.stderr.strip()}")
            if result.returncode != 0:
                raise Exception(f"AppleScript失敗: {result.stderr.strip()[:600]}")

            safe_date = invoice_date.replace("年","").replace("月","").replace("日","")
            filename  = f"{customer_no}_請求書_{safe_date}.numbers"
            print(f"  ✅ {filename}")
            self.send_file(tmp_path, filename, "application/x-iwork-numbers-sffnumbers")
        except subprocess.TimeoutExpired:
            print("  ❌ タイムアウト（Numbers起動に150秒以上かかりました）")
            self.send_json(500, {"ok": False, "error": "タイムアウト：Numbers.appの起動に時間がかかりすぎました"})
        except Exception as e:
            print(f"  ❌ Numbers生成エラー: {e}")
            self.send_json(500, {"ok": False, "error": str(e)})
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    def handle_invoice_template(self, data):
        """Numbers雛形と同じレイアウトでXLSX請求書をゼロから生成（外部テンプレート不要）"""
        import calendar, re
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

        customer_name = data.get("customerName", "")
        invoice_date  = data.get("invoiceDate", "")
        cases         = data.get("cases", [])
        customer_no   = data.get("customerNo", "")

        # 支払い期日：次月末
        due_date = ""
        m = re.match(r"(\d+)年(\d+)月(\d+)日", invoice_date)
        if m:
            y, mo = int(m.group(1)), int(m.group(2))
            nm = mo + 1 if mo < 12 else 1
            ny = y if mo < 12 else y + 1
            due_date = f"{ny}年{nm}月{calendar.monthrange(ny, nm)[1]}日"

        content_parts = [case_label(c) for c in cases if case_label(c)]
        content_str = "、".join(content_parts) or "動画編集案件の件"
        total = sum(int(float(c.get("amount") or c.get("price") or 0)) for c in cases)

        # ---- ワークブック作成 ----
        wb = Workbook()
        ws = wb.active
        ws.title = "請求書"

        # ---- ヘルパー ----
        def S(style='thin'):
            return Side(style=style, color='000000')
        def B(t=None, r=None, b=None, l=None):
            return Border(
                top    = S(t) if t else Side(style=None),
                right  = S(r) if r else Side(style=None),
                bottom = S(b) if b else Side(style=None),
                left   = S(l) if l else Side(style=None),
            )
        def fill(hex_color):
            return PatternFill(fill_type='solid', fgColor=hex_color)

        GRAY  = 'D9D9D9'
        GOLD  = 'FFF2CC'

        # ---- 列幅（Numbers雛形に合わせた比率） ----
        ws.column_dimensions['A'].width = 2
        ws.column_dimensions['B'].width = 32
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 2

        # ---- Row 1: タイトル ----
        ws.row_dimensions[1].height = 36
        ws.merge_cells('B1:C1')
        c = ws['B1']
        c.value     = '請　求　書'
        c.font      = Font(size=22, bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center')

        # ---- Row 2: 宛名 ----
        ws.row_dimensions[2].height = 28
        c = ws['B2']
        c.value     = f'{customer_name}　様'
        c.font      = Font(size=14, bold=True)
        c.alignment = Alignment(vertical='center')
        c.border    = B(b='medium')

        # ---- Rows 3-7: 空白（会社名等スペース） ----
        for r in range(3, 8):
            ws.row_dimensions[r].height = 14

        # ---- Row 8: 請求日 ＋ 「内容」ラベル ----
        ws.row_dimensions[8].height = 18
        c = ws['B8']
        c.value     = f'請求日：{invoice_date}'
        c.font      = Font(size=11)
        c.alignment = Alignment(vertical='center')
        c = ws['C8']
        c.value     = '内　容'
        c.font      = Font(size=10, bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border    = B(t='thin', r='thin', b='thin', l='thin')
        c.fill      = fill(GRAY)

        # ---- Row 9: 支払期日 ＋ 内容テキスト ----
        ws.row_dimensions[9].height = 18
        c = ws['B9']
        c.value     = f'支払い期日：{due_date}'
        c.font      = Font(size=11)
        c.alignment = Alignment(vertical='center')
        c = ws['C9']
        c.value     = content_str
        c.font      = Font(size=10)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border    = B(r='thin', b='thin', l='thin')

        # ---- Rows 10-11: 空白 ----
        for r in [10, 11]:
            ws.row_dimensions[r].height = 10

        # ---- Row 12: 明細ヘッダー ----
        ws.row_dimensions[12].height = 20
        for col, label in [('B', '項　目'), ('C', '金　額（円）')]:
            c = ws[f'{col}12']
            c.value     = label
            c.font      = Font(bold=True, size=10)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border    = B(t='medium', r='thin', b='thin', l='thin')
            c.fill      = fill(GRAY)

        # ---- Rows 13+: 明細データ ----
        for i, case in enumerate(cases):
            r      = 13 + i
            note   = case_label(case)
            amount = int(float(case.get("amount") or case.get("price") or 0))
            ws.row_dimensions[r].height = 18
            bc = ws[f'B{r}']
            bc.value     = note
            bc.font      = Font(size=10)
            bc.alignment = Alignment(vertical='center')
            bc.border    = B(t='thin', r='thin', b='thin', l='thin')
            cc = ws[f'C{r}']
            cc.value         = amount
            cc.font          = Font(size=10)
            cc.number_format = '#,##0'
            cc.alignment     = Alignment(horizontal='right', vertical='center')
            cc.border        = B(t='thin', r='thin', b='thin', l='thin')

        # ---- 小計行 ----
        sr = 13 + len(cases)
        ws.row_dimensions[sr].height = 18
        c = ws[f'B{sr}']
        c.value     = '小　計'
        c.font      = Font(bold=True, size=10)
        c.alignment = Alignment(vertical='center')
        c.border    = B(t='thin', r='thin', b='thin', l='thin')
        c = ws[f'C{sr}']
        c.value         = total
        c.font          = Font(bold=True, size=10)
        c.number_format = '#,##0'
        c.alignment     = Alignment(horizontal='right', vertical='center')
        c.border        = B(t='thin', r='thin', b='thin', l='thin')

        # ---- 消費税・調整欄（Numbers の C16/C17 相当：空白） ----
        for dr in [sr + 1, sr + 2]:
            ws.row_dimensions[dr].height = 18
            ws[f'B{dr}'].border = B(r='thin', b='thin', l='thin')
            ws[f'C{dr}'].border = B(r='thin', b='thin', l='thin')

        # ---- 合計行（Numbers の C18 相当） ----
        tr = sr + 3
        ws.row_dimensions[tr].height = 24
        c = ws[f'B{tr}']
        c.value     = '合　計'
        c.font      = Font(bold=True, size=13)
        c.alignment = Alignment(vertical='center')
        c.border    = B(t='medium', r='thin', b='medium', l='medium')
        c.fill      = fill(GOLD)
        c = ws[f'C{tr}']
        c.value         = total
        c.font          = Font(bold=True, size=13)
        c.number_format = '#,##0'
        c.alignment     = Alignment(horizontal='right', vertical='center')
        c.border        = B(t='medium', r='medium', b='medium', l='thin')
        c.fill          = fill(GOLD)

        # ---- 保存 ----
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        wb.save(tmp.name)

        safe_date = invoice_date.replace("年","").replace("月","").replace("日","")
        filename = f"{customer_no}_請求書_{safe_date}.xlsx"
        print(f"\n📄 XLSX請求書生成: {filename}  ({len(cases)}件 合計{total:,}円)")
        self.send_file(tmp.name, filename,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        os.unlink(tmp.name)

    def handle_add_schedule(self, data):
        """予定を案件表に追加（お客様なしエントリー → Notionカレンダーに同期）"""
        title      = data.get("title","").strip()
        date_str   = data.get("date","")
        memo       = data.get("memo","").strip()
        start_time = data.get("startTime","").strip()  # e.g. "09:00"
        end_time   = data.get("endTime","").strip()    # e.g. "10:00"
        if not title or not date_str:
            self.send_json(400, {"ok": False, "error": "タイトルと日付は必須"}); return

        # 時刻つきの場合はdatetimeフォーマットに変換
        start_dt = f"{date_str}T{start_time}:00+09:00" if start_time else date_str
        end_dt   = f"{date_str}T{end_time}:00+09:00"   if end_time   else None

        date_value = {"start": start_dt}
        if end_dt:
            date_value["end"] = end_dt

        props = {
            "当方案件番号":     {"title": [{"text": {"content": title}}]},
            "案件締切日・進行": {"date": date_value},
            "進捗":            {"status": {"name": "未着手"}},
        }
        if memo:
            props["備考"] = {"rich_text": [{"text": {"content": memo}}]}

        result, err = notion_request("POST", "/pages", {
            "parent": {"database_id": CASE_DB_ID},
            "properties": props,
        })
        if result:
            print(f"  📅 Notion予定追加: 「{title}」{start_dt}")
            self.send_json(200, {"ok": True, "id": result.get("id","")})
        else:
            self.send_json(400, {"ok": False, "error": err})

    def handle_delete_schedule(self, data):
        """予定エントリーをアーカイブ（Notionから削除）"""
        page_id = data.get("id","")
        if not page_id:
            self.send_json(400, {"ok": False, "error": "idが必要"}); return
        # Notion REST API: ページをアーカイブ（= 削除）
        result, err = notion_request("PATCH", f"/pages/{page_id}", {"archived": True})
        if result is not None:
            self.send_json(200, {"ok": True})
        else:
            self.send_json(400, {"ok": False, "error": err})

    def handle_get_all_customers(self):
        """全顧客を取得（顧客情報ツール用）"""
        customers = []
        cursor = None
        while True:
            body = {"sorts": [{"property": "お客様No.", "direction": "ascending"}], "page_size": 100}
            if cursor:
                body["start_cursor"] = cursor
            result, err = notion_request("POST", f"/databases/{CUSTOMER_DB_ID}/query", body)
            if not result:
                self.send_json(400, {"ok": False, "error": err}); return
            for page in result.get("results", []):
                try:
                    props   = page["properties"]
                    no      = (props["お客様No."]["rich_text"] or [{}])[0].get("plain_text","").strip()
                    name    = (props["クライアント名"]["rich_text"] or [{}])[0].get("plain_text","").strip()
                    title_rt= (props["備考"]["title"] or [{}])
                    page_name = title_rt[0].get("plain_text","").strip() if title_rt else ""
                    status  = (props.get("取引状況",{}).get("select") or {}).get("name","")
                    contact = (props.get("お客様優先連絡方法",{}).get("select") or {}).get("name","")
                    kind    = (props.get("種別",{}).get("select") or {}).get("name","")
                    notes_rt= props.get("重要備考",{}).get("rich_text") or []
                    notes   = "".join(t.get("plain_text","") for t in notes_rt)
                    inv_rt  = props.get("請求書宛名",{}).get("rich_text") or []
                    inv_name= "".join(t.get("plain_text","") for t in inv_rt).strip()
                    if no or name:
                        customers.append({
                            "id": page["id"], "no": no, "name": name,
                            "pageName": page_name, "status": status,
                            "contact": contact, "kind": kind, "notes": notes,
                            "invoiceName": inv_name,
                            "notionUrl": page.get("url", ""),
                        })
                except Exception:
                    pass
            if not result.get("has_more"):
                break
            cursor = result.get("next_cursor")
        print(f"\n👥 全顧客取得: {len(customers)}件")
        self.send_json(200, {"ok": True, "customers": customers})

    def handle_bulk_update_customers(self, data):
        """複数顧客の取引状況を一括変更"""
        ids    = data.get("ids", [])
        status = data.get("status", "")
        if not ids or not status:
            self.send_json(400, {"ok": False, "error": "ids と status が必要"}); return
        ok_list, fail_list = [], []
        for pid in ids:
            result, err = notion_request("PATCH", f"/pages/{pid}", {
                "properties": {"取引状況": {"select": {"name": status}}}
            })
            if result:
                ok_list.append(pid)
            else:
                fail_list.append({"id": pid, "error": err})
        self.send_json(200, {"ok": True, "updated": len(ok_list), "failed": len(fail_list), "errors": fail_list})

    def handle_set_invoice_name(self, data):
        """請求書宛名を1件更新する。

        請求書だけ法人名や屋号で出したい、という依頼に対応するための項目。
        空にすればクライアント名で出力される。
        """
        pid  = data.get("id", "")
        name = (data.get("invoiceName") or "").strip()
        if not pid:
            self.send_json(400, {"ok": False, "error": "idが必要"}); return
        rt = [{"text": {"content": name}}] if name else []
        result, err = notion_request("PATCH", f"/pages/{pid}", {
            "properties": {"請求書宛名": {"rich_text": rt}}
        })
        if result:
            print(f"\n🧾 請求書宛名を更新: {name or '（クリア＝クライアント名で出力）'}")
            self.send_json(200, {"ok": True, "invoiceName": name})
        else:
            self.send_json(400, {"ok": False, "error": err})

    def handle_bulk_archive_customers(self, data):
        """複数顧客をアーカイブ（削除）"""
        ids = data.get("ids", [])
        if not ids:
            self.send_json(400, {"ok": False, "error": "ids が必要"}); return
        ok_list, fail_list = [], []
        for pid in ids:
            result, err = notion_request("PATCH", f"/pages/{pid}", {"archived": True})
            if result:
                ok_list.append(pid)
            else:
                fail_list.append({"id": pid, "error": err})
        self.send_json(200, {"ok": True, "archived": len(ok_list), "failed": len(fail_list)})

    def handle_get_customers(self):
        """取引中・頻度低め・取引開始準備中のお客様をNotionから取得"""
        ACTIVE_STATUSES = {"取引中", "頻度低め", "取引開始準備中"}
        customers = []
        cursor = None
        while True:
            body = {"page_size": 100}
            if cursor:
                body["start_cursor"] = cursor
            result, err = notion_request("POST", f"/databases/{CUSTOMER_DB_ID}/query", body)
            if not result:
                self.send_json(400, {"ok": False, "error": err})
                return
            for page in result.get("results", []):
                try:
                    props = page["properties"]
                    no    = (props["お客様No."]["rich_text"] or [{}])[0].get("plain_text","").strip()
                    name  = (props["クライアント名"]["rich_text"] or [{}])[0].get("plain_text","").strip()
                    status = (props.get("取引状況",{}).get("select") or {}).get("name","")
                    contact = (props.get("お客様優先連絡方法",{}).get("select") or {}).get("name","")
                    notes_rt = props.get("重要備考",{}).get("rich_text") or []
                    notes = "".join(t.get("plain_text","") for t in notes_rt)
                    if no and name and status in ACTIVE_STATUSES:
                        # チャンネル情報をパース
                        channels = []
                        for line in notes.split("\n"):
                            line = line.strip()
                            import re
                            m = re.match(r'^([A-Z]{1,2})案件[：:]\s*(.+)', line)
                            if m:
                                channels.append({"lbl": m.group(1), "name": m.group(2).strip()})
                        customers.append({
                            "no": no, "name": name, "status": status,
                            "contact": contact, "channels": channels,
                            "pageId": page["id"],
                            "notionUrl": page.get("url", ""),
                        })
                except Exception:
                    pass
            if not result.get("has_more"):
                break
            cursor = result.get("next_cursor")
        print(f"\n📋 顧客一覧取得: {len(customers)}件")
        self.send_json(200, {"ok": True, "customers": customers})

    def handle_register(self, data):
        print(f"\n📝 案件登録: {data.get('number','?')}  ({data.get('customerNo','?')}様)")

        props = {
            "当方案件番号": {
                "title": [{"text": {"content": data.get("number", "")}}]
            },
            "お客様no/名": {
                "rich_text": [{"text": {"content": data.get("customerNo", "")}}]
            },
            "進捗": {"status": {"name": data.get("progress", "未着手")}},
        }

        if data.get("deadline"):
            props["案件締切日・進行"] = {"date": {"start": data["deadline"]}}
        # 案件名（旧「素材名」）＝ 指定案件ファイル名。
        # 同じプロパティを2箇所から書くと後勝ちで消えるため、優先順を明示する（2026-09-01）
        project_name = (data.get("projectName") or data.get("materialName")
                        or data.get("fileName"))
        if project_name:
            props["指定案件ファイル名"] = {"rich_text": [{"text": {"content": project_name}}]}
        if data.get("memo"):
            props["備考"] = {"rich_text": [{"text": {"content": data["memo"]}}]}
        if data.get("price"):
            try:
                price_val = float(data["price"])
                props["単価"] = {"number": price_val}
                # 外注費がある場合は粗利を計算してテキストで記録
                cost_str = data.get("outsourceCost", "")
                if cost_str:
                    cost_val = float(cost_str)
                    gross = price_val - cost_val
                    def fmt(n):
                        return f"{int(n):,}"
                    gross_text = f"{fmt(price_val)} - {fmt(cost_val)} = {fmt(gross)}"
                    props["粗利（単価-外注費）"] = {"rich_text": [{"text": {"content": gross_text}}]}
            except Exception:
                pass

        result, err = notion_request("POST", "/pages", {
            "parent": {"database_id": CASE_DB_ID},
            "properties": props,
        })

        if result:
            print(f"  ✅ 登録完了: {result.get('url','')}")
            self.send_json(200, {"ok": True, "url": result.get("url", ""), "id": result.get("id", "")})
        else:
            print(f"  ❌ エラー: {err}")
            self.send_json(400, {"ok": False, "error": err})

    def handle_update_customer(self, data):
        no = data.get("customerNo", "")
        entry = data.get("entry", "")
        print(f"\n⭐ 顧客情報追記: {no}番 → {entry}")

        page_id = CUSTOMER_PAGES.get(no)
        if not page_id:
            self.send_json(404, {"ok": False, "error": f"顧客No.{no}のページが見つかりません"})
            return

        # 既存の重要備考を取得
        page, err = notion_request("GET", f"/pages/{page_id}")
        if not page:
            self.send_json(400, {"ok": False, "error": err})
            return

        existing = ""
        try:
            rich = page["properties"]["重要備考"]["rich_text"]
            existing = "".join(t["plain_text"] for t in rich)
        except Exception:
            pass

        new_text = (existing.rstrip() + "\n" + entry).strip()

        result, err = notion_request("PATCH", f"/pages/{page_id}", {
            "properties": {
                "重要備考": {"rich_text": [{"text": {"content": new_text}}]}
            }
        })

        if result:
            print(f"  ✅ 追記完了")
            self.send_json(200, {"ok": True})
        else:
            print(f"  ❌ エラー: {err}")
            self.send_json(400, {"ok": False, "error": err})

    def handle_register_customer(self, data):
        no      = data.get("customerNo", "")
        name    = data.get("name", "")
        print(f"\n👤 顧客登録: {no}番 {name}様")

        props = {
            "備考": {"title": [{"text": {"content": f"{no}_{name}様_補足資料"}}]},
            "お客様No.": {"rich_text": [{"text": {"content": no}}]},
            "クライアント名": {"rich_text": [{"text": {"content": name}}]},
        }
        # 請求書宛名は任意。空ならクライアント名で出力される。
        if (data.get("invoiceName") or "").strip():
            props["請求書宛名"] = {"rich_text": [{"text": {"content": data["invoiceName"].strip()}}]}
        if data.get("type"):
            props["種別"] = {"select": {"name": data["type"]}}
        if data.get("status"):
            props["取引状況"] = {"select": {"name": data["status"]}}
        if data.get("contact"):
            props["お客様優先連絡方法"] = {"select": {"name": data["contact"]}}
        if data.get("notes"):
            props["重要備考"] = {"rich_text": [{"text": {"content": data["notes"]}}]}

        # ① お客様ページを作成
        cust_result, err = notion_request("POST", "/pages", {
            "parent": {"database_id": CUSTOMER_DB_ID},
            "properties": props,
        })
        if not cust_result:
            print(f"  ❌ 顧客登録エラー: {err}")
            self.send_json(400, {"ok": False, "error": err})
            return

        cust_page_id = cust_result["id"]
        cust_url     = cust_result.get("url", "")
        print(f"  ✅ 顧客ページ作成: {cust_url}")

        # 新規顧客を名前キャッシュへ即時反映（案件一覧の名前空欄を防ぐ）
        if no and name:
            CUSTOMER_NAME_CACHE[no] = name

        # ② 補足資料ページを作成（子ページ）
        sub_title = f"{no}_{name}様_補足資料"
        sub_result, sub_err = notion_request("POST", "/pages", {
            "parent": {"page_id": cust_page_id},
            "properties": {
                "title": {"title": [{"text": {"content": sub_title}}]}
            },
            "children": [
                {"object": "block", "type": "heading_2",
                 "heading_2": {"rich_text": [{"text": {"content": "基本情報"}}]}},
                {"object": "block", "type": "paragraph",
                 "paragraph": {"rich_text": [{"text": {"content": f"お客様No.: {no}\nクライアント名: {name}"}}]}},
                {"object": "block", "type": "heading_2",
                 "heading_2": {"rich_text": [{"text": {"content": "案件一覧"}}]}},
                {"object": "block", "type": "paragraph",
                 "paragraph": {"rich_text": [{"text": {"content": "（案件を追記してください）"}}]}},
                {"object": "block", "type": "heading_2",
                 "heading_2": {"rich_text": [{"text": {"content": "連絡先・重要事項"}}]}},
                {"object": "block", "type": "paragraph",
                 "paragraph": {"rich_text": [{"text": {"content": data.get("notes", "（重要事項を追記してください）")}}]}},
            ]
        })

        sub_url = sub_result.get("url", "") if sub_result else ""
        if sub_result:
            print(f"  ✅ 補足資料ページ作成: {sub_url}")
        else:
            print(f"  ⚠️  補足資料ページ作成失敗: {sub_err}")

        self.send_json(200, {
            "ok": True,
            "customerUrl": cust_url,
            "subPageUrl": sub_url,
            "customerNo": no,
            "name": name,
        })


    def handle_get_invoice_data(self):
        """月別案件データをNotionから取得してお客様ごとにグループ化"""
        import re as _re
        from urllib.parse import urlparse, parse_qs
        from datetime import date

        qs = parse_qs(urlparse(self.path).query)
        month_str = qs.get("month", [None])[0]

        today = date.today()
        if month_str:
            try:
                year  = int(month_str.split("-")[0])
                month = int(month_str.split("-")[1])
            except Exception:
                year, month = today.year, today.month
        else:
            # デフォルト: 当月
            year, month = today.year, today.month

        start = f"{year}-{month:02d}-01"
        end   = f"{year+1}-01-01" if month == 12 else f"{year}-{month+1:02d}-01"
        print(f"\n📑 請求データ取得: {year}年{month}月 ({start} 〜 {end})")

        query_body = {
            "filter": {
                "and": [
                    {"property": "案件締切日・進行", "date": {"on_or_after": start}},
                    {"property": "案件締切日・進行", "date": {"before": end}},
                ]
            },
            "sorts": [{"property": "お客様no/名", "direction": "ascending"}],
            "page_size": 100,
        }

        cases = []
        cursor = None
        while True:
            if cursor:
                query_body["start_cursor"] = cursor
            result, err = notion_request("POST", f"/databases/{CASE_DB_ID}/query", query_body)
            if not result:
                self.send_json(400, {"ok": False, "error": err})
                return
            for page in result.get("results", []):
                try:
                    props   = page["properties"]
                    number  = (props["当方案件番号"]["title"] or [{}])[0].get("plain_text","").strip()
                    customer= (props["お客様no/名"]["rich_text"] or [{}])[0].get("plain_text","").strip()
                    price_v = props["単価"]["number"]
                    price   = price_v if price_v is not None else 0
                    note    = (props.get("備考",{}).get("rich_text") or [{}])[0].get("plain_text","").strip()
                    pname   = (props.get("指定案件ファイル名",{}).get("rich_text") or [{}])[0].get("plain_text","").strip()
                    dl      = (props["案件締切日・進行"]["date"] or {}).get("start","")
                    status  = (props["進捗"]["status"] or {}).get("name","")
                    # 最終確認は請求書ツールでチェックするので、案件IDと状態も返す
                    confirmed = bool(props.get("最終確認完了",{}).get("checkbox") or False)
                    if number and customer:
                        cases.append({"id": page["id"],
                                      "number": number, "customer": customer,
                                      "price": price, "note": note,
                                      "projectName": pname, "confirmed": confirmed,
                                      "date": dl, "status": status})
                except Exception as e:
                    pass
            if not result.get("has_more"):
                break
            cursor = result.get("next_cursor")

        # お客様ごとにグループ化
        groups = {}
        for c in cases:
            key = c["customer"]
            if key not in groups:
                groups[key] = {"customer": key, "cases": [], "total": 0}
            groups[key]["cases"].append(c)
            groups[key]["total"] += c["price"]

        # 請求書に載る宛名を添える（請求書宛名 → クライアント名）。
        # 発行前に画面で確認できるようにするため。
        for g in groups.values():
            g["name"] = self.lookup_invoice_name(g["customer"])

        print(f"  → {len(cases)}件 / {len(groups)}お客様")
        self.send_json(200, {
            "ok": True,
            "month": f"{year}-{month:02d}",
            "label": f"{year}年{month}月",
            "customers": list(groups.values()),
        })

    def find_customer_page_id(self, customer_no):
        """お客様NoからNotionページIDを動的検索"""
        if customer_no in CUSTOMER_PAGES:
            return CUSTOMER_PAGES[customer_no]
        body = {
            "filter": {"property": "お客様No.", "rich_text": {"equals": customer_no}},
            "page_size": 1,
        }
        result, _ = notion_request("POST", f"/databases/{CUSTOMER_DB_ID}/query", body)
        if result and result.get("results"):
            pid = result["results"][0]["id"]
            CUSTOMER_PAGES[customer_no] = pid
            return pid
        return None

    def find_invoice_storage_page(self, customer_no):
        """お客様Noの請求書格納庫ページをNotionで検索"""
        # キャッシュキー
        cache_key = f"invoice_{customer_no}"
        if cache_key in CUSTOMER_PAGES:
            return CUSTOMER_PAGES[cache_key]
        result, _ = notion_request("POST", "/search", {
            "query": f"{customer_no}様_請求書格納庫",
            "filter": {"value": "page", "property": "object"},
            "page_size": 10,
        })
        if result:
            for page in result.get("results", []):
                title_rt = (page.get("properties", {}).get("title", {}).get("title") or [])
                title = title_rt[0].get("plain_text", "") if title_rt else ""
                if f"{customer_no}様" in title and "請求書格納庫" in title:
                    pid = page["id"]
                    CUSTOMER_PAGES[cache_key] = pid
                    print(f"  🔍 格納庫発見: {title} ({pid})")
                    return pid
        return None

    def handle_record_invoice(self, data):
        """請求記録をお客様の請求書格納庫ページにトグルで追記"""
        records  = data.get("records", [])
        inv_date = data.get("invoiceDate", "")
        results  = []
        for rec in records:
            cno     = rec.get("customerNo", "")
            amount  = rec.get("amount", 0)
            month   = rec.get("month", "")   # 例: "2026年3月"
            numbers = rec.get("caseNumbers", [])

            # "2026年3月" → "2026年_3月" (格納庫の命名規則に合わせる)
            toggle_title = month.replace("年", "年_") if "年" in month else month

            page_id = self.find_invoice_storage_page(cno)
            if not page_id:
                print(f"  ⚠️  請求書格納庫未発見: {cno}様")
                results.append({"customerNo": cno, "ok": False,
                                 "error": f"{cno}様_請求書格納庫 ページが見つかりません"})
                continue

            detail_text = (f"請求日: {inv_date}　"
                           f"金額: ¥{int(amount):,}　"
                           f"案件 {len(numbers)}件: {', '.join(numbers)}")
            block = {
                "children": [{
                    "object": "block",
                    "type": "toggle",
                    "toggle": {
                        "rich_text": [{"type": "text", "text": {"content": toggle_title}}],
                        "children": [{
                            "object": "block",
                            "type": "paragraph",
                            "paragraph": {
                                "rich_text": [{"type": "text",
                                               "text": {"content": detail_text}}]
                            }
                        }]
                    }
                }]
            }
            res, err = notion_request("PATCH", f"/blocks/{page_id}/children", block)
            if res:
                print(f"  ✅ 格納庫に追記: {cno} → {toggle_title} {detail_text[:50]}")
                results.append({"customerNo": cno, "ok": True, "page": page_id})
            else:
                print(f"  ❌ 追記失敗: {cno} → {err}")
                results.append({"customerNo": cno, "ok": False, "error": err})

        self.send_json(200, {"ok": True, "results": results})


if __name__ == "__main__":
    if NOTION_TOKEN == "secret_ここに貼り付け":
        print("=" * 55)
        print("⚠️  Notionトークンが設定されていません")
        print()
        print("1. https://www.notion.so/my-integrations を開く")
        print("2. 「新しいインテグレーション」をクリック")
        print("3. 名前: 案件登録ツール → 送信")
        print("4. 「シークレット」のトークンをコピー")
        print("5. このファイル（notion_server.py）の")
        print("   NOTION_TOKEN = の行に貼り付けて保存")
        print("=" * 55)
        sys.exit(1)

    import socket
    local_ip = socket.gethostbyname(socket.gethostname())

    print("=" * 55)
    print("📁 案件登録サーバー起動中...")
    print(f"   PC:     http://localhost:{PORT}")
    print(f"   iPhone: http://{local_ip}:{PORT}")
    print("   （同じWiFiで接続してください）")
    print("   停止: Ctrl+C")
    print("=" * 55)

    _threading.Thread(target=eigyo_auto_loop, daemon=True, name="eigyo-auto").start()
    print(f"   🤖 営業メール自動送信: {'ON' if EIGYO_AUTO_DEFAULT else 'OFF'}（{EIGYO_AUTO_DAYS}曜 {EIGYO_AUTO_HOURS}時・1日{min(EIGYO_AUTO_PER_DAY, EIGYO_DAILY_LIMIT)}通まで・{EIGYO_AUTO_GAP_SEC}秒間隔{'・テスト送信' if EIGYO_AUTO_DRY else ''}）")
    server = HTTPServer(("0.0.0.0", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n✅ サーバーを停止しました")
