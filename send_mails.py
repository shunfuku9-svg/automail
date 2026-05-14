from __future__ import annotations

import argparse
import os
import re
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
MESSAGE_PATH = BASE_DIR / "message.txt"
ENV_PATH = BASE_DIR / ".env"
DEFAULT_SUBJECT = "ご連絡のお願い | 松本高専 電気電子工学科"


@dataclass
class Recipient:
    domain: str
    kanji_name: str
    romaji_name: str
    email: str
    company: str


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gmail メール送信（対話モード）")
    parser.add_argument("--message", default=str(MESSAGE_PATH), help="本文テンプレート")
    parser.add_argument(
        "--subject",
        default=os.environ.get("MAIL_SUBJECT", DEFAULT_SUBJECT),
        help="件名",
    )
    return parser.parse_args()


def is_valid_email(email: str) -> bool:
    return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email.strip()))


def normalize_name_parts(name: str) -> tuple[str, str]:
    raw = name.strip()
    if not raw:
        return "", ""
    cleaned = re.sub(r"[()\/,\u3001]", " ", raw)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    parts = [re.sub(r"[^a-z0-9-]", "", piece.lower()) for piece in re.split(r"[ \u3000]+", cleaned)]
    parts = [piece for piece in parts if piece]
    if len(parts) >= 2:
        return parts[0], parts[-1]
    if len(parts) == 1:
        return parts[0], ""
    return "", ""


def extract_domains(value: str) -> list[str]:
    raw = value.strip().lower()
    if not raw:
        return []
    domains: list[str] = []
    seen: set[str] = set()
    for token in re.split(r"[,;\s]+", raw):
        if not token:
            continue
        if "@" in token:
            token = token.split("@", 1)[1]
        token = re.sub(r"^https?://", "", token)
        token = token.split("/", 1)[0]
        token = re.sub(r"^www\.", "", token)
        if "." not in token:
            token = token + ".com"
        if re.fullmatch(r"[a-z0-9.-]+\.[a-z]{2,}", token) and token not in seen:
            seen.add(token)
            domains.append(token)
    return domains


def create_email_candidates(name: str, domain_input: str) -> list[str]:
    first, last = normalize_name_parts(name)
    domains = extract_domains(domain_input)
    if not domains:
        return []
    patterns: list[str] = []
    if first and last:
        fi, li = first[:1], last[:1]
        patterns.extend(
            [
                f"{first}.{last}", f"{first}{last}", f"{first}_{last}", f"{first}-{last}",
                f"{fi}{last}", f"{fi}.{last}", f"{fi}_{last}", f"{fi}-{last}",
                f"{first}{li}", f"{first}.{li}", f"{first}_{li}", f"{first}-{li}",
                f"{last}.{first}", f"{last}{first}", f"{last}_{first}", f"{last}-{first}",
                f"{last}{fi}", f"{last}.{fi}", first, last,
                f"{fi}{li}", f"{li}{fi}",
            ]
        )
    elif first:
        patterns.extend([first, first[:1]])
    seen: set[str] = set()
    emails: list[str] = []
    for pattern in patterns:
        if not pattern:
            continue
        for domain in domains:
            email = f"{pattern}@{domain}"
            if email not in seen:
                seen.add(email)
                emails.append(email)
    return emails


def render_template(template: str, recipient: Recipient) -> str:
    values = {
        "company": recipient.company or "ご担当者",
        "kanji_name": recipient.kanji_name or "ご担当者",
        "romaji_name": recipient.romaji_name or "",
        "email": recipient.email or "",
        "domain": recipient.domain or "",
    }
    output = template
    for key, value in values.items():
        output = output.replace("{" + key + "}", value)
    return output


def send_email(sender: str, app_password: str, recipient: Recipient, subject: str, body: str) -> None:
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient.email
    msg["Subject"] = subject
    msg.set_content(body)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender, app_password)
        smtp.send_message(msg)


def prompt(question: str, default: str = "") -> str:
    if default:
        s = input(f"{question} [{default}]: ").strip()
        return s if s else default
    return input(f"{question}: ").strip()


def confirm(question: str, default: bool = False) -> bool:
    suffix = " [Y/n]" if default else " [y/N]"
    ans = input(f"{question}{suffix}: ").strip().lower()
    if not ans:
        return default
    return ans in ("y", "yes", "はい")


def main() -> int:
    load_env_file(ENV_PATH)
    args = parse_args()

    print("\n=== メール送信ウィザード ===\n")

    company = prompt("会社名を入力してください")
    domain = prompt("宛先のドメイン（または URL）を入力してください")
    kanji_name = prompt("宛名（漢字）を入力してください")
    romaji_name = prompt("ローマ字名を入力してください（候補メール生成用、Enter でスキップ）")

    direct_email = prompt("送信先メールアドレスを直接指定しますか？（Enter で候補から自動生成）")

    candidate_name = romaji_name or kanji_name
    if direct_email and is_valid_email(direct_email):
        emails_to_send = [direct_email.strip()]
    else:
        candidates = create_email_candidates(candidate_name, domain)
        if not candidates:
            print("候補メールアドレスを生成できませんでした。ドメインとローマ字名を確認してください。")
            return 1
        if len(candidates) == 1:
            emails_to_send = candidates
        else:
            print("\n候補メールアドレス:")
            for i, addr in enumerate(candidates, 1):
                print(f"  {i}: {addr}")
            choice = prompt(f"番号で選択（1-{len(candidates)}）、または 0 ですべて送信", "1")
            try:
                n = int(choice)
                if n == 0:
                    emails_to_send = candidates
                elif 1 <= n <= len(candidates):
                    emails_to_send = [candidates[n - 1]]
                else:
                    emails_to_send = [candidates[0]]
            except ValueError:
                emails_to_send = [candidates[0]]

    recipient = Recipient(
        domain=domain,
        kanji_name=kanji_name,
        romaji_name=romaji_name,
        email="",
        company=company,
    )

    message_path = Path(args.message)
    message_template = message_path.read_text(encoding="utf-8")
    body = render_template(message_template, recipient)

    print("\n--- プレビュー ---")
    print(f"宛先: {', '.join(emails_to_send)}")
    print(f"件名: {args.subject}")
    print(f"本文:\n{body}")
    print("---\n")

    if not confirm("この内容で送信しますか？"):
        print("キャンセルしました。")
        return 0

    sender = os.environ.get("GMAIL_ADDRESS", "").strip()
    app_password = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "").strip()
    if not sender or not app_password:
        print("エラー: .env に GMAIL_ADDRESS と GMAIL_APP_PASSWORD を設定してください。")
        return 1

    for email in emails_to_send:
        recipient.email = email
        send_email(sender, app_password, recipient, args.subject, body)
        print(f"送信完了: {email}")

    print("\n送信が完了しました。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
