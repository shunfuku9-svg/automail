# automail

Gmail メール送信ツール（対話モード）。会社名・ドメイン・宛名を一問一答で入力して送信します。

## 機能

- 会社ドメイン＋ローマ字名から候補メールアドレスを自動生成
- `message.txt` の `{company}` `{kanji_name}` を置換して本文生成
- Gmail SMTP で送信

## 事前準備

`.env` を作成し、Gmail の**アプリ パスワード**を設定してください。

```env
GMAIL_ADDRESS=your-address@gmail.com
GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
```

件名は `MAIL_SUBJECT` で上書きできます。

## 実行方法

```bash
python3 send_mails.py
```

起動後、以下の順で入力します。

1. 会社名
2. ドメイン（または URL）
3. 宛名（漢字）
4. ローマ字名（候補メール生成用、Enter でスキップ）
5. 送信先メールアドレス（直接指定する場合。Enter で候補から自動生成）
6. 候補が複数ある場合は番号で選択、または 0 ですべて送信
7. プレビュー確認後に送信の可否を選択

## オプション

| オプション | 説明 |
|-----------|------|
| `--message パス` | 本文テンプレート（デフォルト: message.txt） |
| `--subject "件名"` | 件名を指定 |

## ファイル

| ファイル | 説明 |
|---------|------|
| `send_mails.py` | メインスクリプト |
| `message.txt` | 本文テンプレート。{company} {kanji_name} は置換されます |
| `.env` | Gmail 認証情報（gitignore 済み） |
| `.env.example` | .env のサンプル |
