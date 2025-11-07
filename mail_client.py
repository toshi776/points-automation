#!/usr/bin/env python3
"""Yahooメール(IMAP/SMTP)の疎通確認とメール概要取得を行うユーティリティ."""

from __future__ import annotations

import argparse
import base64
import imaplib
import os
import smtplib
import ssl
from dataclasses import dataclass
from email import message_from_bytes
from email.header import decode_header, Header
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, List, Optional


ENV_PATH = Path(__file__).resolve().parent / ".env"
DEFAULT_IMAP_HOST = "imap.mail.yahoo.co.jp"
DEFAULT_IMAP_PORT = 993
DEFAULT_SMTP_HOST = "smtp.mail.yahoo.co.jp"
DEFAULT_SMTP_PORT = 465


def load_env_file(path: Path = ENV_PATH) -> None:
    """Lightweight .env loader (コメントと空行のみ無視)."""
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def ensure_env(var_names: Iterable[str]) -> None:
    missing = [name for name in var_names if not os.getenv(name)]
    if missing:
        joined = ", ".join(missing)
        raise RuntimeError(f"環境変数 {joined} が設定されていません。.env かシェルで定義してください。")


def decode_mime_header(value: Optional[bytes | str]) -> str:
    if value in (None, "", b""):
        return ""

    if isinstance(value, bytes):
        try:
            value = value.decode("utf-8")
        except UnicodeDecodeError:
            value = value.decode("iso-8859-1", errors="replace")

    header = decode_header(value)
    decoded_parts = []
    for text, encoding in header:
        if isinstance(text, bytes):
            decoded_parts.append(text.decode(encoding or "utf-8", errors="replace"))
        else:
            decoded_parts.append(text)
    return "".join(decoded_parts)


@dataclass
class MailConfig:
    login_id: str
    mail_address: str
    password: str
    imap_host: str = DEFAULT_IMAP_HOST
    imap_port: int = DEFAULT_IMAP_PORT
    smtp_host: str = DEFAULT_SMTP_HOST
    smtp_port: int = DEFAULT_SMTP_PORT


def load_config() -> MailConfig:
    ensure_env(["PASSWORD"])
    login_id = os.getenv("ID") or os.getenv("MAIL_ADDRESS")
    mail_address = os.getenv("MAIL_ADDRESS") or login_id
    if not login_id or not mail_address:
        raise RuntimeError("ID もしくは MAIL_ADDRESS が必要です。.env を確認してください。")

    return MailConfig(
        login_id=login_id,
        mail_address=mail_address,
        password=os.environ["PASSWORD"],
        imap_host=os.getenv("IMAP_HOST", DEFAULT_IMAP_HOST),
        imap_port=int(os.getenv("IMAP_PORT", DEFAULT_IMAP_PORT)),
        smtp_host=os.getenv("SMTP_HOST", DEFAULT_SMTP_HOST),
        smtp_port=int(os.getenv("SMTP_PORT", DEFAULT_SMTP_PORT)),
    )


def encode_mailbox_name(name: str) -> str:
    """IMAPフォルダ名をASCIIに収まるよう Modified UTF-7 へ変換."""
    if not name:
        return name
    try:
        name.encode("ascii")
        return name
    except UnicodeEncodeError:
        encoded_chunks = []
        buffer = []

        def flush_buffer():
            if not buffer:
                return
            chunk = "".join(buffer)
            utf16 = chunk.encode("utf-16-be")
            b64 = base64.b64encode(utf16).decode("ascii").replace("/", ",").rstrip("=")
            encoded_chunks.append(f"&{b64}-")
            buffer.clear()

        for char in name:
            code = ord(char)
            if char == "&":
                flush_buffer()
                encoded_chunks.append("&-")
            elif 0x20 <= code <= 0x7E:
                flush_buffer()
                encoded_chunks.append(char)
            else:
                buffer.append(char)

        flush_buffer()
        return "".join(encoded_chunks)


def fetch_recent_messages(config: MailConfig, folder: str, limit: int) -> List[dict]:
    context = ssl.create_default_context()
    with imaplib.IMAP4_SSL(config.imap_host, config.imap_port, ssl_context=context) as imap:
        imap.login(config.login_id, config.password)
        mailbox = encode_mailbox_name(folder)
        imap.select(mailbox)
        status, data = imap.search(None, "ALL")
        if status != "OK":
            raise RuntimeError(f"{folder} の検索に失敗しました: {status}")

        message_ids = data[0].split()
        if not message_ids:
            return []

        limited_ids = message_ids[-limit:]
        summaries = []
        for msg_id in reversed(limited_ids):
            status, msg_data = imap.fetch(msg_id, "(BODY.PEEK[HEADER])")
            if status != "OK":
                continue

            raw_email = msg_data[0][1]
            message = message_from_bytes(raw_email)
            summaries.append(
                {
                    "id": msg_id.decode(),
                    "subject": decode_mime_header(message.get("Subject")),
                    "from": decode_mime_header(message.get("From")),
                    "date": message.get("Date", ""),
                }
            )
        return summaries


def send_test_email(config: MailConfig, to_address: str, subject: str, body: str) -> None:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = config.mail_address
    msg["To"] = to_address
    msg.set_content(body)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(config.smtp_host, config.smtp_port, context=context) as smtp:
        smtp.login(config.login_id, config.password)
        smtp.send_message(msg)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Yahooメール IMAP/SMTP 疎通ユーティリティ")
    sub = parser.add_subparsers(dest="command")

    fetch_parser = sub.add_parser("fetch", help="指定フォルダから最新メールの概要を取得")
    fetch_parser.add_argument("--folder", default="INBOX", help="IMAPフォルダ（デフォルト: INBOX）")
    fetch_parser.add_argument("--limit", type=int, default=5, help="取得件数（最新から）")

    send_parser = sub.add_parser("send-test", help="SMTPでテストメールを送信")
    send_parser.add_argument("--to", help="送信先メールアドレス。未指定時は MAIL_ADDRESS")
    send_parser.add_argument("--subject", default="Points Automation SMTP Test", help="テストメール件名")
    send_parser.add_argument(
        "--body",
        default="Yahoo SMTP送信テスト\nこのメールは points-automation スクリプトから送信されました。",
        help="テストメール本文",
    )
    return parser


def main() -> None:
    load_env_file()
    parser = build_parser()
    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return

    config = load_config()
    if args.command == "fetch":
        messages = fetch_recent_messages(config, folder=args.folder, limit=max(args.limit, 1))
        if not messages:
            print(f"{args.folder} にはメールがありませんでした。")
            return

        for msg in messages:
            print("=" * 60)
            print(f"ID    : {msg['id']}")
            print(f"Date  : {msg['date']}")
            print(f"From  : {msg['from']}")
            print(f"Subj  : {msg['subject']}")
        return

    if args.command == "send-test":
        to_addr = args.to or config.mail_address
        send_test_email(config, to_address=to_addr, subject=args.subject, body=args.body)
        print(f"SMTP送信完了: {config.mail_address} -> {to_addr}")
        return


if __name__ == "__main__":
    main()
