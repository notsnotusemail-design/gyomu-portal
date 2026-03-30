#!/usr/bin/env python3
# coding: utf-8
"""
案件登録サーバー
NotionのAPIへ直接登録します。
起動: python3 notion_server.py
"""

import sys, json, os, tempfile, datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.request import urlopen, Request
from urllib.error import HTTPError, URLError
import json as jsonlib
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

# 日次スケジュール保存ファイル（ローカルJSON）
DAILY_SCHEDULE_FILE = os.path.join(SCRIPT_DIR, "daily_schedules.json")

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
        elif path == "/api/health":
            self.send_json(200, {"status": "ok", "message": "サーバー起動中"})
        elif path == "/api/next-customer-no":
            nos = get_all_customer_nos()
            self.send_json(200, {
                "ok": True,
                "next_300s":    next_in_range(nos, 300, 399),
                "next_3000s":   next_in_range(nos, 3000, 3999),
                "next_regular": next_in_range(nos, 1, 99),
            })
        elif path == "/api/customers":
            self.handle_get_customers()
        elif path == "/api/customers-all":
            self.handle_get_all_customers()
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
        elif path.startswith("/api/daily-schedule/dates"):
            self.handle_get_schedule_dates()
        elif path.startswith("/api/daily-schedule"):
            self.handle_get_daily_schedule()
        elif path.startswith("/api/handover"):
            self.handle_get_handover()
        else:
            self.send_json(404, {"error": "Not found"})

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        try:
            data = jsonlib.loads(raw)
        except Exception:
            self.send_json(400, {"error": "Invalid JSON"})
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
        else:
            self.send_json(404, {"error": "Not found"})

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

    def _generate_invoice_inner(self, data):
        customer_no   = data.get('customerNo', '')
        invoice_date  = data.get('invoiceDate', '')
        cases         = data.get('cases', [])

        # お客様名をNotionから取得（案件表のお客様no/名フィールド値で検索）
        customer_name = data.get('customerName', '')
        if not customer_name:
            body = {"filter": {"property": "お客様No.", "rich_text": {"equals": customer_no}}, "page_size": 1}
            res, _ = notion_request("POST", f"/databases/{CUSTOMER_DB_ID}/query", body)
            if res and res.get("results"):
                props = res["results"][0].get("properties", {})
                name_rt = props.get("クライアント名", {}).get("rich_text") or []
                customer_name = name_rt[0].get("plain_text", "") if name_rt else ""
            if not customer_name:
                customer_name = customer_no + "様"
        if not customer_name.endswith("様"):
            customer_name += "様"

        ORANGE = "B07C1A"
        BLUE   = "4472C4"

        def S(): return Side(border_style='thin', color='000000')
        def B(t=False, b=False, l=False, r=False):
            return Border(top=S() if t else Side(), bottom=S() if b else Side(),
                          left=S() if l else Side(), right=S() if r else Side())

        wb = Workbook()
        ws = wb.active
        ws.title = "請求書"
        ws.column_dimensions['A'].width = 1.5
        ws.column_dimensions['B'].width = 36
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 18

        def c(row, col, val=None, font=None, align=None, border=None, fmt=None):
            cell = ws.cell(row, col)
            if val is not None:  cell.value = val
            if font:   cell.font = font
            if align:  cell.alignment = align
            if border: cell.border = border
            if fmt:    cell.number_format = fmt
            return cell

        r = 1
        c(r,2,'請求書 No.', font=Font(name='メイリオ',size=8,color=BLUE)); r+=1
        ws.row_dimensions[r].height = 34
        c(r,2, customer_name, font=Font(name='メイリオ',size=22,color=ORANGE),
          align=Alignment(vertical='center'))
        ws.merge_cells(f'B{r}:C{r}'); r+=1
        ws.row_dimensions[r].height = 6; r+=1
        ws.row_dimensions[r].height = 6; r+=1
        c(r,2,'野津　欧', font=Font(name='メイリオ',size=11,color=ORANGE)); r+=1
        c(r,2,'三菱UFJ銀行新宿通り支店　（050）-0571808',
          font=Font(name='メイリオ',size=9,color=ORANGE)); r+=1
        ws.row_dimensions[r].height = 10; r+=1

        ws.row_dimensions[r].height = 20
        c(r,2, f'請求日：{invoice_date}',
          font=Font(name='メイリオ',size=10,color=ORANGE), border=B(t=True,l=True))
        c(r,3,'内容', font=Font(name='メイリオ',size=10,color=ORANGE,bold=True),
          align=Alignment(horizontal='center',vertical='center'), border=B(t=True,l=True,r=True)); r+=1

        ws.row_dimensions[r].height = 20
        c(r,2,'支払い期日：', font=Font(name='メイリオ',size=10), border=B(b=True,l=True))
        c(r,3,'動画編集案件の件', font=Font(name='メイリオ',size=10), border=B(b=True,l=True,r=True)); r+=1
        ws.row_dimensions[r].height = 8; r+=1

        ws.row_dimensions[r].height = 20
        c(r,2,'詳細', font=Font(name='メイリオ',size=10,color=ORANGE,bold=True), border=B(t=True,b=True,l=True))
        c(r,3,'金額', font=Font(name='メイリオ',size=10,color=ORANGE,bold=True),
          align=Alignment(horizontal='right',vertical='center'), border=B(t=True,b=True,l=True,r=True))
        c(r,4,'当方案件番号', font=Font(name='メイリオ',size=9,color=ORANGE)); r+=1

        total = 0
        for case in cases:
            desc    = case.get('note') or case.get('number','')
            amount  = int(case.get('amount') or case.get('price') or 0)
            case_no = case.get('number','')
            total  += amount
            ws.row_dimensions[r].height = 18
            c(r,2,desc, font=Font(name='メイリオ',size=10),
              align=Alignment(wrap_text=True,vertical='center'), border=B(b=True,l=True))
            c(r,3,amount, font=Font(name='メイリオ',size=10),
              align=Alignment(horizontal='right',vertical='center'),
              border=B(b=True,l=True,r=True), fmt='#,##0')
            c(r,4,case_no, font=Font(name='メイリオ',size=9,color='AAAAAA')); r+=1

        ws.row_dimensions[r].height = 8; r+=1

        for text, bold, bordered in [
            (f'小計　¥{total:,}',  False, False),
            ('税率　0%',           False, False),
            ('その他　¥0',         False, False),
            (f'集計　¥{total:,}',  True,  True),
        ]:
            ws.merge_cells(f'B{r}:C{r}')
            c(r,2,text, font=Font(name='メイリオ',size=10,color=ORANGE,bold=bold),
              align=Alignment(horizontal='center'),
              border=B(t=True,b=True,l=True,r=True) if bordered else None); r+=1

        ws.row_dimensions[r].height = 12; r+=1
        c(r,2,'この請求書に関してご不明な点がございましたら、お問い合わせください。',
          font=Font(name='メイリオ',size=9,color=ORANGE)); r+=2
        c(r,2,'今月もありがとうございます',
          font=Font(name='メイリオ',size=11,color=ORANGE,bold=True))

        ws.print_area = f'B1:C{r+1}'
        ws.page_margins.left = 0; ws.page_margins.right = 0
        ws.page_margins.top  = 0.4; ws.page_margins.bottom = 0.4

        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        wb.save(tmp.name)
        safe = invoice_date.replace('/', '-')
        filename = f'{customer_no}_請求書_{safe}.xlsx'
        print(f"\n📄 請求書生成: {filename}")
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

    def handle_add_schedule(self, data):
        """予定を案件表に追加（お客様なしエントリー → Notionカレンダーに同期）"""
        title    = data.get("title","").strip()
        date_str = data.get("date","")
        memo     = data.get("memo","").strip()
        if not title or not date_str:
            self.send_json(400, {"ok": False, "error": "タイトルと日付は必須"}); return

        props = {
            "当方案件番号":     {"title": [{"text": {"content": title}}]},
            "案件締切日・進行": {"date":  {"start": date_str}},
            "進捗":            {"status": {"name": "未着手"}},
        }
        if memo:
            props["備考"] = {"rich_text": [{"text": {"content": memo}}]}

        result, err = notion_request("POST", "/pages", {
            "parent": {"database_id": CASE_DB_ID},
            "properties": props,
        })
        if result:
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
                    if no or name:
                        customers.append({
                            "id": page["id"], "no": no, "name": name,
                            "pageName": page_name, "status": status,
                            "contact": contact, "kind": kind, "notes": notes,
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
        if data.get("materialName"):
            props["備考/素材名"] = {"rich_text": [{"text": {"content": data["materialName"]}}]}
        if data.get("fileName"):
            props["指定案件ファイル名"] = {"rich_text": [{"text": {"content": data["fileName"]}}]}
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
                    note    = (props["備考/素材名"]["rich_text"] or [{}])[0].get("plain_text","").strip()
                    dl      = (props["案件締切日・進行"]["date"] or {}).get("start","")
                    status  = (props["進捗"]["status"] or {}).get("name","")
                    if number and customer:
                        cases.append({"number": number, "customer": customer,
                                      "price": price, "note": note,
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

    server = HTTPServer(("0.0.0.0", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n✅ サーバーを停止しました")
