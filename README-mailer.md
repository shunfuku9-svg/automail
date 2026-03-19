# Local Mailer

VS Code terminal based mailer with optional Claude drafting.

## Files

- `message.txt`: base draft template
- `ai-prompt.txt`: Claude drafting instructions
- `recipients.csv`: recipient data
- `send-mails.ps1`: generator + drafter + sender

## CSV columns

- `domain`: company domain used for candidate generation
- `kanji_name`: recipient name for the email body
- `romaji_name`: recipient name used for email candidate generation
- `emails`: optional explicit email addresses; if empty, candidates are generated
- `status`: `送信済み` or `sent` will be skipped
- `company`: company name
- `sent_at`: filled after sending

## Environment variable

Set your Anthropic API key before using `-UseClaude`.

```powershell
$env:ANTHROPIC_API_KEY="your_api_key_here"
```

## Commands

Show generated email candidates only:

```powershell
powershell -ExecutionPolicy Bypass -File .\send-mails.ps1 -GenerateOnly
```

Dry run without OpenAI:

```powershell
powershell -ExecutionPolicy Bypass -File .\send-mails.ps1 -DryRun
```

Dry run with Claude drafting:

```powershell
powershell -ExecutionPolicy Bypass -File .\send-mails.ps1 -DryRun -UseClaude -Model claude-sonnet-4-20250514
```

Send with Claude drafting:

```powershell
powershell -ExecutionPolicy Bypass -File .\send-mails.ps1 -UseClaude -Model claude-sonnet-4-20250514 -Subject "【ご相談のお願い】東京高専 吉本舜一"
```

## Flow

1. Edit `message.txt`
2. Edit `recipients.csv`
3. Optionally edit `ai-prompt.txt`
4. Run `-GenerateOnly`
5. Run `-DryRun` or `-DryRun -UseClaude`
6. Send

## Notes

- Gmail sending still uses your Gmail address and Gmail App Password.
- Claude only drafts the body; Gmail handles delivery.
- Generated candidate emails may be wrong, so test with your own address first.
