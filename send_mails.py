from __future__ import annotations

import argparse
import csv
import os
import re
import smtplib
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Dict, List


BASE_DIR = Path(__file__).resolve().parent
CSV_PATH = BASE_DIR / "recipients.csv"
MESSAGE_PATH = BASE_DIR / "message.txt"
PREVIEW_PATH = BASE_DIR / "preview.txt"
ENV_PATH = BASE_DIR / ".env"

STATUS_SENT = "送信済み"
STATUS_SKIP = "スキップ"
STATUS_REVIEW = "要確認"
STATUS_NO_CANDIDATE = "候補なし"
STATUS_SENT_ALL = "全候補送信済み"
DEFAULT_SUBJECT = "ご連絡のお願い | 松本高専 電気電子工学科"
AUTO_SEND_FIRST_CANDIDATE = True


@dataclass
class Recipient:
    domain: str
    kanji_name: str
    romaji_name: str
    email: str
    status: str
    company: str
    sent_at: str
    row: Dict[str, str]


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
    parser = argparse.ArgumentParser(description="Gmail auto mailer")
    parser.add_argument("--csv", default=str(CSV_PATH))
    parser.add_argument("--message", default=str(MESSAGE_PATH))
    parser.add_argument("--preview", default=str(PREVIEW_PATH))
    parser.add_argument("--subject", default=os.environ.get("MAIL_SUBJECT", DEFAULT_SUBJECT))
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--generate-only", action="store_true")
    parser.add_argument("--review-only", action="store_true")
    return parser.parse_args()


def load_recipients(csv_path: Path) -> tuple[List[Recipient], List[str]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        fieldnames = list(reader.fieldnames or [])
        required = ["domain", "kanji_name", "email", "status", "company"]
        for name in required:
            if name not in fieldnames:
                raise ValueError(f"CSV に必須列がありません: {name}")
        if "romaji_name" not in fieldnames:
            fieldnames.insert(2, "romaji_name")
        if "sent_at" not in fieldnames:
            fieldnames.append("sent_at")

        recipients: List[Recipient] = []
        for row in reader:
            normalized = {key: (value or "").strip() for key, value in row.items()}
            normalized.setdefault("sent_at", "")
            recipients.append(
                Recipient(
                    domain=normalized.get("domain", ""),
                    kanji_name=normalized.get("kanji_name", ""),
                    romaji_name=normalized.get("romaji_name", ""),
                    email=normalized.get("email", ""),
                    status=normalized.get("status", ""),
                    company=normalized.get("company", ""),
                    sent_at=normalized.get("sent_at", ""),
                    row=normalized,
                )
            )
    return recipients, fieldnames


def save_recipients(csv_path: Path, fieldnames: List[str], recipients: List[Recipient]) -> None:
    with csv_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for recipient in recipients:
            data = dict(recipient.row)
            data["domain"] = recipient.domain
            data["kanji_name"] = recipient.kanji_name
            data["romaji_name"] = recipient.romaji_name
            data["email"] = recipient.email
            data["status"] = recipient.status
            data["company"] = recipient.company
            data["sent_at"] = recipient.sent_at
            writer.writerow(data)


def is_valid_email(email: str) -> bool:
    return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email.strip()))


def split_emails(value: str) -> List[str]:
    emails: List[str] = []
    seen = set()
    for token in re.split(r"[,;\s]+", value.strip()):
        if not token:
            continue
        if is_valid_email(token) and token not in seen:
            seen.add(token)
            emails.append(token)
    return emails


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


def extract_domains(value: str) -> List[str]:
    raw = value.strip().lower()
    if not raw:
        return []

    domains: List[str] = []
    seen = set()
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


def create_email_candidates(name: str, domain_input: str) -> List[str]:
    first, last = normalize_name_parts(name)
    domains = extract_domains(domain_input)
    if not domains:
        return []

    patterns: List[str] = []
    if first and last:
        first_initial = first[:1]
        last_initial = last[:1]
        patterns.extend(
            [
                f"{first}.{last}",
                f"{first}{last}",
                f"{first}_{last}",
                f"{first}-{last}",
                f"{first_initial}{last}",
                f"{first_initial}.{last}",
                f"{first_initial}_{last}",
                f"{first_initial}-{last}",
                f"{first}{last_initial}",
                f"{first}.{last_initial}",
                f"{first}_{last_initial}",
                f"{first}-{last_initial}",
                f"{last}.{first}",
                f"{last}{first}",
                f"{last}_{first}",
                f"{last}-{first}",
                f"{last}{first_initial}",
                f"{last}.{first_initial}",
                first,
                last,
                f"{first_initial}{last_initial}",
                f"{last_initial}{first_initial}",
            ]
        )
    elif first:
        patterns.extend([first, first[:1]])

    seen = set()
    emails: List[str] = []
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


def write_preview(preview_path: Path, recipients: List[Recipient], message_template: str) -> None:
    blocks: List[str] = []
    for recipient in recipients:
        candidate_name = recipient.romaji_name or recipient.kanji_name
        candidates = split_emails(recipient.email) or create_email_candidates(candidate_name, recipient.domain)
        blocks.append(
            "\n".join(
                [
                    f"NAME: {recipient.kanji_name}",
                    f"ROMAJI: {recipient.romaji_name or '(none)'}",
                    "EMAILS: " + (", ".join(candidates) if candidates else "(none)"),
                    "",
                    render_template(message_template, recipient),
                ]
            )
        )
    preview_text = ("\n\n" + ("-" * 40) + "\n\n").join(blocks)
    preview_path.write_text(preview_text, encoding="utf-8")


def create_message(sender: str, recipient: Recipient, subject: str, body: str) -> EmailMessage:
    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient.email
    message["Subject"] = subject
    message.set_content(body)
    return message


def send_email(sender: str, app_password: str, recipient: Recipient, subject: str, body: str) -> None:
    message = create_message(sender, recipient, subject, body)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender, app_password)
        smtp.send_message(message)


def main() -> int:
    load_env_file(ENV_PATH)
    args = parse_args()

    csv_path = Path(args.csv)
    message_path = Path(args.message)
    preview_path = Path(args.preview)

    recipients, fieldnames = load_recipients(csv_path)
    message_template = message_path.read_text(encoding="utf-8")

    write_preview(preview_path, recipients, message_template)

    summary = {"sent": 0, "review": 0, "skip": 0, "no_candidate": 0}

    if args.generate_only:
        for recipient in recipients:
            if split_emails(recipient.email) or not recipient.kanji_name:
                continue
            candidate_name = recipient.romaji_name or recipient.kanji_name
            candidates = create_email_candidates(candidate_name, recipient.domain)
            if candidates:
                recipient.email = ", ".join(candidates)
                recipient.status = STATUS_REVIEW
                summary["review"] += 1
            else:
                recipient.status = STATUS_NO_CANDIDATE
                summary["no_candidate"] += 1
        save_recipients(csv_path, fieldnames, recipients)
        print(f"候補生成完了 review={summary['review']} no_candidate={summary['no_candidate']}")
        return 0

    sender = os.environ.get("GMAIL_ADDRESS", "").strip()
    app_password = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "").strip()
    if not args.dry_run and (not sender or not app_password):
        raise ValueError(".env または環境変数に GMAIL_ADDRESS / GMAIL_APP_PASSWORD を設定してください")

    for recipient in recipients:
        if recipient.status.strip() in {STATUS_SENT, STATUS_SENT_ALL}:
            continue
        if not recipient.kanji_name:
            recipient.status = STATUS_SKIP
            summary["skip"] += 1
            continue
        emails_to_send = split_emails(recipient.email)
        if not emails_to_send:
            candidate_name = recipient.romaji_name or recipient.kanji_name
            candidates = create_email_candidates(candidate_name, recipient.domain)
            if not candidates:
                recipient.status = STATUS_NO_CANDIDATE
                summary["no_candidate"] += 1
                continue
            recipient.email = ", ".join(candidates)
            recipient.status = STATUS_REVIEW
            summary["review"] += 1
            if args.review_only or not AUTO_SEND_FIRST_CANDIDATE:
                continue
            emails_to_send = candidates

        if not emails_to_send:
            recipient.status = STATUS_NO_CANDIDATE
            summary["no_candidate"] += 1
            continue

        body = render_template(message_template, recipient)
        if args.dry_run:
            for email in emails_to_send:
                print(f"[DRY RUN] to={email} subject={args.subject}")
        else:
            for email in emails_to_send:
                recipient.email = email
                send_email(sender, app_password, recipient, args.subject, body)
            recipient.email = ", ".join(emails_to_send)
            recipient.status = STATUS_SENT_ALL if len(emails_to_send) > 1 else STATUS_SENT
            recipient.sent_at = datetime.now().isoformat(timespec="seconds")
            summary["sent"] += len(emails_to_send)

    save_recipients(csv_path, fieldnames, recipients)
    print(
        "完了 "
        f"sent={summary['sent']} review={summary['review']} "
        f"skip={summary['skip']} no_candidate={summary['no_candidate']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
