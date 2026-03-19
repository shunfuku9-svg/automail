# Gmail 自動送信システム

このフォルダで完結する Python 版のメール自動送信ツールです。

## 機能

- `recipients.csv` を読み込んで送信対象を管理
- 会社ドメインから候補メールアドレスを生成
- 候補メールアドレスを `email` 列へすべて書き込む
- 送信前にプレビューを保存
- `--dry-run` で誤送信せずに確認
- Gmail SMTP で実送信
- 送信後に `status` と `sent_at` を更新
- 候補アドレスを自動生成した行もそのまま送信可能

## ファイル

- `send_mails.py`
  メインスクリプト
- `run_mailer.ps1`
  Python の場所を自動判定して実行するラッパー
- `recipients.csv`
  宛先一覧
- `message.txt`
  本文テンプレート
- `.env.example`
  Gmail 設定の例
- `preview.txt`
  直近プレビューの出力先

## CSV 列

- `domain`
  企業ドメイン、URL、またはメールアドレス断片
- `kanji_name`
  宛名に使う氏名
- `romaji_name`
  候補メールアドレス生成に使うローマ字名
- `email`
  実際に送るメールアドレス。複数ある場合はカンマ区切り
- `status`
  `送信済み`、`要確認`、`スキップ`、`候補なし` など
- `company`
  会社名
- `sent_at`
  送信日時

## Gmail 設定

Gmail の通常パスワードではなく、Google アカウントのアプリ パスワードを使います。

`.env` を作って以下を設定してください。

```env
GMAIL_ADDRESS=your-address@gmail.com
GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
```

## 使い方

候補メールアドレスと本文を確認するだけ:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1 --generate-only
```

送信せずにプレビューだけ更新:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1 --dry-run
```

実際に送信:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1
```

候補を作るだけで止めたい場合:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1 --review-only
```

件名を上書き:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_mailer.ps1 --subject "ご連絡のお願い | 松本高専 電気電子工学科"
```

## 注意

- デフォルトでは、候補メールアドレスを自動生成した行は `email` 列へ全候補を書き込み、その全件へ送信します
- 候補生成を使う場合は `romaji_name` を入れてください。漢字氏名だけでは推測できません
- 候補生成だけで止めたいときは `--review-only` を使ってください
- 最初は自分の Gmail 宛てだけで試してください
- `python` や `py` が PATH に無くても `run_mailer.ps1` から実行できます
