#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""案件表の「備考/素材名」に残った内容を「備考」へ統合して、旧項目を空にする。

案件名は migrate_案件名.py で「指定案件ファイル名」へ移し終えている前提。
ここでは移せずに残っている分（衝突などで保留した行）を備考へ寄せて、
備考の二重管理を解消する。既定はドライラン。--apply で実行。
"""
import os, sys, json, time, datetime, urllib.request, urllib.error

TOKEN  = os.environ.get("NOTION_TOKEN", "")
DB_ID  = "2e13288b84b081faa8e8fb22589ce378"
SRC    = "備考/素材名"
DST    = "備考"

def req(method, path, body=None):
    r = urllib.request.Request("https://api.notion.com/v1" + path,
        data=json.dumps(body).encode() if body else None, method=method,
        headers={"Authorization": f"Bearer {TOKEN}", "Notion-Version": "2022-06-28",
                 "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(r) as f: return json.load(f)
    except urllib.error.HTTPError as e:
        print(f"  ! APIエラー {e.code}: {e.read().decode()[:200]}"); return None

def rt(props, name):
    return "".join(t.get("plain_text", "") for t in (props.get(name) or {}).get("rich_text") or [])

def title(props, name):
    return "".join(t.get("plain_text", "") for t in (props.get(name) or {}).get("title") or [])

def fetch_all():
    out, cur = [], None
    while True:
        body = {"page_size": 100}
        if cur: body["start_cursor"] = cur
        res = req("POST", f"/databases/{DB_ID}/query", body)
        if not res: sys.exit("案件表の取得に失敗しました")
        out += res["results"]
        if not res.get("has_more"): return out
        cur = res["next_cursor"]

def main():
    if not TOKEN: sys.exit("NOTION_TOKEN が未設定です")
    apply_ = "--apply" in sys.argv
    print(f"モード: {'★本番実行' if apply_ else 'ドライラン（書き換えなし）'}\n")

    pages = fetch_all()
    print(f"案件表 {len(pages)} 件を取得\n")

    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    with open(f"backup_備考統合_{stamp}.json", "w", encoding="utf-8") as f:
        json.dump([{"id": p["id"], "案件番号": title(p["properties"], "当方案件番号"),
                    SRC: rt(p["properties"], SRC), DST: rt(p["properties"], DST)}
                   for p in pages], f, ensure_ascii=False, indent=1)
    print(f"現状を backup_備考統合_{stamp}.json に保存しました\n")

    jobs, skip = [], 0
    for p in pages:
        pr  = p["properties"]
        src = rt(pr, SRC).strip()
        if not src:
            continue
        dst = rt(pr, DST).strip()
        # 既に備考へ書かれている内容は足さない（重複を増やさない）
        merged = dst if src in dst else (f"{dst} / {src}" if dst else src)
        if merged == dst: skip += 1
        jobs.append((p["id"], title(pr, "当方案件番号"), dst, src, merged))

    print(f"■ 素材名に残っている行 : {len(jobs)} 件（うち備考に既出 {skip} 件は備考を変更せず消去のみ）\n")
    for _, no, dst, src, merged in jobs:
        print(f"  {no:20s} 備考『{dst[:24]}』＋素材名『{src[:24]}』")
        print(f"  {'':20s}   → 備考『{merged[:60]}』")

    if not apply_:
        print("\nドライランのため書き換えていません。実行するには --apply を付けてください。")
        return

    print("\n書き換えを開始します…")
    ok = ng = 0
    for pid, no, dst, src, merged in jobs:
        res = req("PATCH", f"/pages/{pid}", {"properties": {
            DST: {"rich_text": [{"text": {"content": merged}}]},
            SRC: {"rich_text": []},
        }})
        if res: ok += 1
        else:   ng += 1; print(f"  失敗: {no}")
        time.sleep(0.35)
    print(f"\n完了: 成功 {ok} 件 / 失敗 {ng} 件")

if __name__ == "__main__":
    main()
