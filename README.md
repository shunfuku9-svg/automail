# automail

Gmail 自動送信システム。`recipients.csv` を読み込み、メールを自動送信します。

## 機能

- `recipients.csv` を読み込んで送信対象を管理
- 会社ドメインから候補メールアドレスを生成
- 候補メールアドレスを `email` 列へすべて書き込む
- 送信前にプレビューを保存
- `--dry-run` で誤送信せずに確認
- Gmail SMTP で実送信
- 送信後に `status` と `sent_at` を更新

## 事前準備

### Gmail 設定（`.env`）

`.env` を作成し、以下を設定してください。**通常パスワードではなく、Google アカウントのアプリ パスワード**を使います。

```env
GMAIL_ADDRESS=your-address@gmail.com
GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
```

### 主要ファイル

- `send_mails.py` — メインスクリプト
- `recipients.csv` — 宛先一覧
- `message.txt` — 本文テンプレート
- `preview.txt` — 直近プレビューの出力先

## 実行方法

### 対話モード（おすすめ）

会社名・ドメイン・宛名などを一問一答で入力して1件ずつ送信します。

```bash
python send_mails.py --interactive
# または
python send_mails.py -i
```

起動後、以下の順で入力していきます。

1. 会社名
2. ドメイン（または URL）
3. 宛名（漢字）
4. ローマ字名（候補メール生成用）
5. 送信先メールアドレス（直接指定する場合。Enter で候補から自動生成）
6. 候補が複数ある場合は番号で選択、または 0 ですべて送信
7. プレビュー確認後に送信の可否を選択

### バッチモード（CSV 一括）

### macOS / Linux

```bash
# 候補メールアドレスと本文を確認するだけ
python send_mails.py --generate-only

# 送信せずにプレビューだけ更新
python send_mails.py --dry-run

# 実際に送信
python send_mails.py

# 候補を生成して止める（確認してから送信）
python send_mails.py --review-only

# 件名を指定して送信
python send_mails.py --subject "ご連絡のお願い | 松本高専 電気電子工学科"
```

### Windows

```powershell
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1 --generate-only
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1 --dry-run
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1 --subject "件名"
```

## オプション一覧

| オプション | 説明 |
|-----------|------|
| `-i`, `--interactive` | 対話モード（一問一答で1件ずつ送信） |
| `--generate-only` | 候補メールアドレスと本文を確認 |
| `--dry-run` | 送信せずにプレビューを `preview.txt` に出力 |
| `--review-only` | 候補を生成して停止（送信しない） |
| `--subject "件名"` | メールの件名を指定 |

## 注意事項

- 最初は `--dry-run` で動作確認してから、実際に送信してください
- 最初は自分の Gmail 宛てだけで試すことを推奨します
- 候補生成を使う場合は `romaji_name` を CSV に入れてください（漢字氏名だけでは推測できません）
