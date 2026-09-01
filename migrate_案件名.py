#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""案件表のプロパティ再編（素材名 → 指定案件ファイル名 へ転記、備考を一本化）

背景:
  「備考/素材名」という項目が実態と合っていない。ここに入っているのは
  クライアント指定の回名なので、本来の「指定案件ファイル名（＝指定案件名）」へ移す。
  空いた「備考/素材名」は廃止し、備考は既存の「備考」1本に統合する。

やること（1レコードずつ）:
  1. 「備考/素材名」の値を「指定案件ファイル名」へ転記
     - 転記先が空 → そのまま書き込む
     - 転記先に既に値があり、内容が同じ → 何もしない
     - 転記先に既に値があり、内容が違う → ★衝突として保留（上書きしない）
  2. 転記できたレコードは「備考/素材名」を空にする（--clear 指定時のみ）

使い方:
  export NOTION_TOKEN=ntn_xxxxx
  python3 migrate_案件名.py --inspect      # ★まずこれ。項目一覧と実データを確認
  python3 migrate_案件名.py                # ドライラン（何も書き換えない）
  python3 migrate_案件名.py --apply        # 転記を実行
  python3 migrate_案件名.py --apply --clear # 転記＋素材名の消去まで実行

安全策:
  - 既定はドライラン。--apply を付けない限り 1件も書き換えない
  - 実行前に全レコードの現状を backup_案件名_<日時>.json に保存する
  - 衝突レコードは絶対に上書きせず、一覧で報告する
"""
import os, sys, json, time, datetime
import urllib.request, urllib.error

TOKEN = os.environ.get("NOTION_TOKEN", "")
CASE_DB_ID = "2e13288b-84b0-81fa-a8e8-fb22589ce378"
API = "https://api.notion.com/v1"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
}

SRC_PROP  = "備考/素材名"          # 移行元（廃止予定）
DST_PROP  = "指定案件ファイル名"    # 移行先（案件名）※--dst で変更可
MEMO_PROP = "備考"                # 統合後の唯一の備考

# --dst で移行先プロパティ名を上書きできる（例: --dst 指定案件名）
for _i, _a in enumerate(sys.argv):
    if _a == "--dst" and _i + 1 < len(sys.argv):
        DST_PROP = sys.argv[_i + 1]

def req(method, path, body=None):
    data = json.dumps(body).encode() if body is not None else None
    r = urllib.request.Request(API + path, data=data, headers=HEADERS, method=method)
    try:
        with urllib.request.urlopen(r) as res:
            return json.loads(res.read().decode())
    except urllib.error.HTTPError as e:
        print(f"  ! APIエラー {e.code}: {e.read().decode()[:200]}")
        return None

def rt(props, name):
    """rich_text プロパティの平文を取り出す"""
    p = props.get(name) or {}
    arr = p.get("rich_text") or []
    return (arr[0].get("plain_text", "") if arr else "").strip()

def title(props, name):
    p = props.get(name) or {}
    arr = p.get("title") or []
    return (arr[0].get("plain_text", "") if arr else "").strip()

def fetch_all():
    pages, cursor = [], None
    while True:
        body = {"page_size": 100}
        if cursor: body["start_cursor"] = cursor
        res = req("POST", f"/databases/{CASE_DB_ID}/query", body)
        if not res: sys.exit("案件表の取得に失敗しました")
        pages += res.get("results", [])
        if not res.get("has_more"): break
        cursor = res.get("next_cursor")
    return pages

def inspect():
    """DBのプロパティ一覧と、サンプル案件の実データを表示して事実確認する"""
    db = req("GET", f"/databases/{CASE_DB_ID}")
    if not db: sys.exit("DBスキーマの取得に失敗しました")
    print("=== 案件表のプロパティ一覧 ===")
    for name, meta in db.get("properties", {}).items():
        print(f"  {name:24s} : {meta.get('type')}")
    print()
    sample = "301-26-0331-B"
    res = req("POST", f"/databases/{CASE_DB_ID}/query",
              {"filter": {"property": "当方案件番号", "title": {"equals": sample}}, "page_size": 1})
    if res and res.get("results"):
        pr = res["results"][0]["properties"]
        print(f"=== サンプル案件 {sample} の実データ ===")
        for name, meta in pr.items():
            t = meta.get("type")
            if t == "title":      v = title(pr, name)
            elif t == "rich_text": v = rt(pr, name)
            elif t == "number":    v = meta.get("number")
            elif t == "status":    v = (meta.get("status") or {}).get("name")
            elif t == "date":      v = (meta.get("date") or {}).get("start")
            else:                  v = f"<{t}>"
            if v not in (None, "", "<>"):
                print(f"  {name:24s} = {v}")
        print()
        print("★『そうのすけくん』が入っている項目名を確認してください。")
        print("   それが案件名の移行先です。異なる場合は --dst 項目名 を付けて実行してください。")
    else:
        print(f"サンプル案件 {sample} が見つかりませんでした")

def main():
    if not TOKEN:
        sys.exit("NOTION_TOKEN が未設定です。\n  export NOTION_TOKEN=ntn_xxxxx  を実行してから再度お試しください。")
    if "--inspect" in sys.argv:
        inspect(); return
    apply_ = "--apply" in sys.argv
    clear_ = "--clear" in sys.argv
    print(f"移行先プロパティ: 「{DST_PROP}」")

    print(f"モード: {'★本番実行' if apply_ else 'ドライラン（書き換えなし）'}"
          f"{' ＋ 素材名を消去' if (apply_ and clear_) else ''}\n")

    pages = fetch_all()
    print(f"案件表 {len(pages)} 件を取得\n")

    # 現状バックアップ
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = [{"id": p["id"],
               "案件番号": title(p["properties"], "当方案件番号"),
               SRC_PROP: rt(p["properties"], SRC_PROP),
               DST_PROP: rt(p["properties"], DST_PROP),
               MEMO_PROP: rt(p["properties"], MEMO_PROP)} for p in pages]
    bkfile = f"backup_案件名_{stamp}.json"
    with open(bkfile, "w", encoding="utf-8") as f:
        json.dump(backup, f, ensure_ascii=False, indent=1)
    print(f"現状を {bkfile} に保存しました\n")

    move, same, empty, conflict = [], [], [], []
    same_ids = []
    for p in pages:
        pr = p["properties"]
        no  = title(pr, "当方案件番号")
        src = rt(pr, SRC_PROP)
        dst = rt(pr, DST_PROP)
        if not src:            empty.append((no, dst))
        elif not dst:          move.append((p["id"], no, src))
        elif src == dst:       same.append((no, src)); same_ids.append((p["id"], no))
        else:                  conflict.append((no, src, dst))

    print(f"■ 転記する            : {len(move)} 件")
    print(f"■ 既に同じ値（--clear時は素材名だけ消去）: {len(same)} 件")
    print(f"■ 素材名が空（対象外）  : {len(empty)} 件")
    print(f"■ ★衝突（上書きしない） : {len(conflict)} 件\n")

    if move:
        print("--- 転記内容（先頭20件）---")
        for _, no, src in move[:20]:
            print(f"  {no:20s} {SRC_PROP}『{src[:40]}』 → {DST_PROP}")
        if len(move) > 20: print(f"  …ほか {len(move)-20} 件")
        print()
    if conflict:
        print("--- ★衝突：両方に値があり内容が違うため保留 ---")
        for no, src, dst in conflict:
            print(f"  {no:20s} 素材名『{src[:32]}』 / 既存ファイル名『{dst[:32]}』")
        print("  → どちらを残すか決めてから個別に対応してください\n")

    if not apply_:
        print("ドライランのため書き換えていません。実行するには --apply を付けてください。")
        return

    # --clear のときは「既に同じ値」の行も対象にする。
    # 値が完全に一致しているので消しても情報は失われない。
    # 衝突している行はどちらを残すか判断が要るため手を付けない。
    jobs = [(pid, no, src, True) for pid, no, src in move]
    if clear_:
        jobs += [(pid, no, None, False) for pid, no in same_ids]

    print(f"書き換えを開始します…（{len(jobs)} 件）")
    ok = ng = 0
    for pid, no, src, write_dst in jobs:
        props = {}
        if write_dst:
            props[DST_PROP] = {"rich_text": [{"text": {"content": src}}]}
        if clear_:
            props[SRC_PROP] = {"rich_text": []}
        res = req("PATCH", f"/pages/{pid}", {"properties": props})
        if res: ok += 1
        else:   ng += 1; print(f"  失敗: {no}")
        time.sleep(0.35)   # Notion APIのレート制限対策（約3req/秒）
    print(f"\n完了: 成功 {ok} 件 / 失敗 {ng} 件")
    if not clear_:
        print("※ 素材名の消去は行っていません。転記結果を確認後 --clear を付けて再実行してください。")

if __name__ == "__main__":
    main()
