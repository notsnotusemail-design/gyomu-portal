# 営業メール 自動送信 — 設定手順

## 1. GitHub に上げる（2ファイル）
- notion_server.py
- 営業ツール.html
（GitHub の Web 画面「Add files via upload」で上書き → Railway が自動で再デプロイ）

## 2. Railway の Variables に入れる
| 変数 | 値の例 | 意味 |
|---|---|---|
| GMAIL_USER | notsnotusemail@gmail.com | 送信元Gmail（既にあれば不要） |
| GMAIL_APP_PASSWORD | xxxx xxxx xxxx xxxx | Gmailのアプリパスワード（既にあれば不要） |
| EIGYO_SENDER_ADDRESS | 〒xxx-xxxx 東京都…（事業所の住所） | **必須**。特定電子メール法で本文に住所が要る |
| EIGYO_SENDER_TEL | 080-xxxx-xxxx | 任意（テンプレに無いときだけ末尾に付く） |
| EIGYO_SENDER_URL | https://… | 任意 |
| EIGYO_AUTO | 0 | 起動時にONにするなら 1（画面のボタンでもON/OFFできる） |
| EIGYO_AUTO_HOURS | 10-18 | 送る時間帯（日本時間） |
| EIGYO_AUTO_DAYS | 1-5 | 送る曜日（1=月 … 7=日） |
| EIGYO_AUTO_PER_DAY | 40 | 自動で送る1日の上限（EIGYO_DAILY_LIMIT 以下） |
| EIGYO_AUTO_GAP_SEC | 480 | 送信の間隔（秒）。480=8分おき |
| EIGYO_LOW_WATER | 20 | 残りがこれを切ったら引き継ぎに「営業リストの補充」を書く |
| EIGYO_AUTO_DRY | 0 | 1にすると送ったことにして実際は送らない（最初の動作確認用） |

## 3. 最初の動かし方（おすすめ）
1. EIGYO_AUTO_DRY=1 で再デプロイ → 営業ツール「自動送信」タブ →「今すぐ1通送る」で本文を確認（送られない・リストも変わらない）
2. 問題なければ EIGYO_AUTO_DRY を 0 に戻す → 「自動送信をONにする」
3. 以後は時間帯の中で8分おきに1通ずつ送られ、送った会社は「送信済」になる

## 4. 止め方・注意
- 「OFFにする」ボタン、または Railway で EIGYO_AUTO=0 にして再デプロイ
- Gmail のログインに失敗したら自動でOFFになり、記録に理由が出る
- 送れない会社（宛先不正・同じ会社に送信済 など）は「送信失敗」になり、メモの先頭に理由が入る
- 「不要」と返信が来た会社は、手で「営業お断り」にする（以後は送らない）
- 対象は **メールアドレスがある未着手** だけ。フォームしか無い会社は今は対象外（次の段階でフォーム送信を足す）

---

## ★ 追記（2026-10-27）Railway では SMTP が使えない → Gmail API で送る

Railway の中から smtp.gmail.com（465/587）につながらず OSError になる（コンソールで実測）。
HTTPS は通るので、**Gmail API** で送る方式に切り替えた。アプリパスワードは不要になる。

### 1. Google Cloud で OAuth クライアントを作る（1回だけ・10分）
1. https://console.cloud.google.com/ → プロジェクトを新規作成（名前は何でも。例「営業ツール」）
2. 「APIとサービス」→「ライブラリ」→ **Gmail API** を検索 → 有効にする
3. 「APIとサービス」→「OAuth 同意画面」→ 外部 → アプリ名「営業ツール」、メールは自分 → 保存
   - スコープは追加しなくてよい（スクリプト側で gmail.send を要求する）
   - テストユーザーに自分の Gmail を追加
   - **最後に「アプリを公開」→ 本番環境に公開**（テストのままだとトークンが7日で切れる。「確認」は不要、未確認の警告が出るだけ）
4. 「認証情報」→「認証情報を作成」→「OAuth クライアント ID」→ 種類 **デスクトップ アプリ** → 作成
   → 出てきた **クライアントID** と **クライアントシークレット** を控える

### 2. 手元のMacでリフレッシュトークンを作る
```
python3 ~/Documents/業務ポータル_アップロード_営業自動送信/gmail_oauth_setup.py
```
クライアントID・シークレットを貼る → ブラウザで Gmail にログインして「許可」
（「このアプリは Google で確認されていません」→「詳細」→「営業ツール（安全ではないページ）に移動」で進む）
→ ターミナルに3行出る

### 3. Railway の Variables に3つ入れてデプロイ
| 変数 | 値 |
|---|---|
| GMAIL_OAUTH_CLIENT_ID | 2.で出た値 |
| GMAIL_OAUTH_CLIENT_SECRET | 2.で出た値 |
| GMAIL_OAUTH_REFRESH_TOKEN | 2.で出た値 |

GMAIL_APP_PASSWORD は残しておいても消してもよい（Gmail API の設定があればそちらが優先）。

### 4. 確認
営業ツール →「自動送信」タブ → 「送り主」の横に「Gmail API」と出ていれば OK → 「今すぐ1通送る」→ 記録に「送信」が出る → 「自動送信をONにする」
