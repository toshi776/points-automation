#!/usr/bin/env python3
"""Yahooメール(IMAP/SMTP)の疎通確認とメール概要取得を行うユーティリティ."""

from __future__ import annotations

import argparse
import base64
import imaplib
import os
import re
import smtplib
import ssl
import textwrap
import json
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email import message_from_bytes
from email.header import decode_header, Header
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
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


def parse_email_date(value: str) -> Optional[datetime]:
    if not value:
        return None
    try:
        parsed = parsedate_to_datetime(value)
    except (TypeError, ValueError):
        return None
    if parsed is None:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def _strip_html_tags(html: str) -> str:
    # Best-effort removal of HTML tags for keyword検出.
    clean = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", html)
    clean = re.sub(r"<[^>]+>", " ", clean)
    clean = re.sub(r"\s+", " ", clean)
    return clean.strip()


def extract_plain_text(message) -> str:
    parts: List[str] = []
    html_fallback: List[str] = []
    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            if content_type.startswith("multipart/"):
                continue
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace")
            if content_type == "text/plain":
                parts.append(text)
            elif content_type == "text/html":
                html_fallback.append(_strip_html_tags(text))
    else:
        payload = message.get_payload(decode=True) or b""
        charset = message.get_content_charset() or "utf-8"
        text = payload.decode(charset, errors="replace")
        if message.get_content_type() == "text/html":
            html_fallback.append(_strip_html_tags(text))
        else:
            parts.append(text)

    if parts:
        return "\n".join(parts)
    if html_fallback:
        return "\n".join(html_fallback)
    return ""


LIMIT_PATTERNS = [
    r"期間限定",
    r"期間中",
    r"今だけ",
    r"本日まで",
    r"本日?限り",
    r"\d{1,2}月\d{1,2}日まで",
    r"\d{1,2}/\d{1,2}まで",
    r"\d+日以内",
    r"締切",
    r"過去最高",
    r"超還元",
    r"緊急",
]

POINT_PATTERNS = [
    r"ポイント.?アップ",
    r"ポイント.{0,4}倍",
    r"ポイント増量",
    r"ポイント還元",
    r"ボーナスポイント",
    r"ポイント加算",
    r"還元",
    r"\d{3,}\s*P",
    r"Pが.*?倍",
]


def detect_limited_point_campaign(text: str) -> tuple[bool, List[str], List[str]]:
    normalized = text.replace("\u3000", " ")
    limit_hits = [pat for pat in LIMIT_PATTERNS if re.search(pat, normalized)]
    point_hits = [pat for pat in POINT_PATTERNS if re.search(pat, normalized)]
    return bool(limit_hits and point_hits), limit_hits, point_hits


GENRE_PATTERNS = {
    "会員登録": [
        r"会員登録",
        r"無料登録",
        r"メンバー登録",
        r"入会",
    ],
    "口座": [
        r"口座開設",
        r"口座申込",
        r"銀行口座",
    ],
    "クレカ": [
        r"カード発行",
        r"クレジットカード",
        r"年会費",
        r"タッチ決済",
    ],
    "FX": [
        r"FX",
        r"lot",
        r"外為",
        r"為替",
    ],
    "カードローン": [
        r"カードローン",
        r"キャッシング",
        r"融資",
    ],
    "サブスク": [
        r"サブスク",
        r"定額サービス",
        r"月額",
        r"ストリーミング",
    ],
    "光回線": [
        r"光回線",
        r"光インターネット",
        r"フレッツ",
        r"回線工事",
    ],
    "でんき乗り換え": [
        r"でんき",
        r"電力",
        r"電気料金",
        r"エネルギー",
    ],
    "不動産": [
        r"不動産",
        r"査定",
        r"訪問査定",
        r"マンション",
        r"住宅",
        r"投資用物件",
    ],
    "保険": [
        r"保険",
        r"見積",
        r"保険料",
        r"損保",
        r"生保",
    ],
    "投資・証券": [
        r"証券",
        r"株",
        r"NISA",
        r"投資",
        r"資産運用",
        r"ロボアド",
    ],
    "医療・美容": [
        r"クリニック",
        r"エステ",
        r"医療",
        r"美容",
        r"脱毛",
        r"カウンセリング",
    ],
    "住設サービス": [
        r"ウォーターサーバー",
        r"太陽光",
        r"ソーラー",
        r"リフォーム",
        r"水回り",
    ],
    "法人サービス": [
        r"法人",
        r"SaaS",
        r"ビジネスローン",
        r"B2B",
        r"資料請求",
    ],
}

ALLOWED_GENRES = list(GENRE_PATTERNS.keys())
HIGH_VALUE_GENRES = {
    "口座",
    "クレカ",
    "FX",
    "カードローン",
    "不動産",
    "投資・証券",
    "保険",
    "医療・美容",
    "住設サービス",
    "法人サービス",
}


def detect_genres(subject: str, body: str) -> List[str]:
    text = f"{subject}\n{body}".lower()
    detected: List[str] = []
    seen = set()
    for genre, patterns in GENRE_PATTERNS.items():
        for pat in patterns:
            if re.search(pat.lower(), text):
                if genre not in seen:
                    detected.append(genre)
                    seen.add(genre)
                break
    return detected


POINT_VALUE_PATTERN = re.compile(r"(\d[\d,]{2,})\s*P", re.IGNORECASE)


def extract_point_values(text: str) -> List[int]:
    values: List[int] = []
    for match in POINT_VALUE_PATTERN.finditer(text):
        raw = match.group(1).replace(",", "")
        try:
            values.append(int(raw))
        except ValueError:
            continue
    return values


HIGH_VALUE_KEYWORDS = re.compile(
    r"(口座開設|口座申込|カード発行|個別面談|面談|査定|訪問査定|訪問見積|契約|ローン申し込み|カウンセリング|相談会)"
)


def should_override_purchase_filter(subject: str, body: str, genres: List[str]) -> bool:
    text = f"{subject}\n{body}"
    values = extract_point_values(text)
    max_value = max(values) if values else 0
    if max_value < 3000:
        return False
    if HIGH_VALUE_KEYWORDS.search(text):
        return True
    return False


SECTION_HEADER_PATTERN = re.compile(r"(?m)^◆.+◆\s*$")


def split_campaign_blocks(body: str) -> List[str]:
    text = body.replace("\r\n", "\n")
    matches = list(SECTION_HEADER_PATTERN.finditer(text))
    if not matches:
        stripped = text.strip()
        return [stripped] if stripped else []
    blocks: List[str] = []
    for idx, match in enumerate(matches):
        start = match.start()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        segment = text[start:end].strip()
        if segment:
            blocks.append(segment)
    if not blocks:
        stripped = text.strip()
        return [stripped] if stripped else []
    return blocks


def derive_block_subject(global_subject: str, block_text: str, block_index: int) -> str:
    for line in block_text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("◆") or stripped.startswith("▲"):
            continue
        if stripped.lower().startswith("https://") or stripped.lower().startswith("http://"):
            continue
        if stripped.lower().startswith("pr"):
            continue
        return stripped
    return f"{global_subject} (part {block_index + 1})"


def has_point_keyword(text: str) -> bool:
    if re.search(r"\d{2,}\s*P", text):
        return True
    if re.search(r"P\s*\d{2,}", text):
        return True
    if re.search(r"\d.{0,6}ポイント", text):
        return True
    if re.search(r"ポイント.{0,6}\d", text):
        return True
    if re.search(r"ポイント(加算|還元|付与|増量|アップ|UP)", text):
        return True
    return False


def is_percentage_only(text: str) -> bool:
    percent = re.search(r"\d+(?:\.\d+)?\s*%.*?(GET|還元|バック)", text, flags=re.IGNORECASE)
    if not percent:
        return False
    return not has_point_keyword(text)


def is_purchase_percent_campaign(text: str) -> bool:
    if not re.search(r"\d+(?:\.\d+)?\s*%.*?(GET|還元|キャッシュバック)", text, flags=re.IGNORECASE):
        return False
    if not re.search(r"購入|商品|ショッピング|決済|支払|利用", text):
        return False
    return True


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


def fetch_messages_since(config: MailConfig, folder: str, days: int) -> List[dict]:
    context = ssl.create_default_context()
    since_dt = datetime.now(timezone.utc) - timedelta(days=days)
    since_str = since_dt.strftime("%d-%b-%Y")
    with imaplib.IMAP4_SSL(config.imap_host, config.imap_port, ssl_context=context) as imap:
        imap.login(config.login_id, config.password)
        mailbox = encode_mailbox_name(folder)
        imap.select(mailbox)
        status, data = imap.search(None, "SINCE", since_str)
        if status != "OK":
            raise RuntimeError(f"{folder} の検索に失敗しました: {status}")

        message_ids = data[0].split()
        if not message_ids:
            return []

        summaries: List[dict] = []
        for msg_id in reversed(message_ids):
            status, msg_data = imap.fetch(msg_id, "(RFC822)")
            if status != "OK" or not msg_data:
                continue
            raw_email = msg_data[0][1]
            message = message_from_bytes(raw_email)
            plain_text = extract_plain_text(message)
            parsed_date = parse_email_date(message.get("Date", ""))
            summaries.append(
                {
                    "id": msg_id.decode(),
                    "subject": decode_mime_header(message.get("Subject")),
                    "from": decode_mime_header(message.get("From")),
                    "date": message.get("Date", ""),
                    "received_at": parsed_date,
                    "body": plain_text,
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


def find_point_campaigns(messages: List[dict]) -> tuple[List[dict], List[dict]]:
    campaigns: List[dict] = []
    rejected: List[dict] = []
    for msg in messages:
        body = msg.get("body") or ""
        email_subject = msg.get("subject", "")
        blocks = split_campaign_blocks(body)
        if not blocks:
            blocks = [body]
        for block_index, block_text in enumerate(blocks):
            block_subject = derive_block_subject(email_subject, block_text, block_index)
            genres = [g for g in detect_genres(block_subject, block_text) if g in ALLOWED_GENRES]
            override_purchase = should_override_purchase_filter(block_subject, block_text, genres)
            matched, limit_hits, point_hits = detect_limited_point_campaign(block_text)
            reasons: List[str] = []
            if not matched and not override_purchase:
                reasons.append("keyword_not_match")
            if matched or override_purchase:
                if is_percentage_only(block_text) and not override_purchase:
                    reasons.append("percent_only")
                if is_purchase_percent_campaign(block_text) and not override_purchase:
                    reasons.append("purchase_percent")
            if not genres:
                reasons.append("genre_none")
            if reasons:
                rejected.append(
                    {
                        "id": msg.get("id"),
                        "block_index": block_index,
                        "email_subject": email_subject,
                        "subject": block_subject,
                        "from": msg.get("from"),
                        "date": msg.get("date"),
                        "reasons": reasons,
                    }
                )
                continue
            snippet = textwrap.shorten(" ".join(block_text.split()), width=160, placeholder="…")
            campaigns.append(
                {
                    "id": msg.get("id"),
                    "block_index": block_index,
                    "email_subject": email_subject,
                    "subject": block_subject,
                    "from": msg.get("from"),
                    "date": msg.get("date"),
                    "matched_limit": limit_hits,
                    "matched_point": point_hits,
                    "genres": genres,
                    "snippet": snippet,
                }
            )
    return campaigns, rejected


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Yahooメール IMAP/SMTP 疎通ユーティリティ")
    sub = parser.add_subparsers(dest="command")

    fetch_parser = sub.add_parser("fetch", help="指定フォルダから最新メールの概要を取得")
    fetch_parser.add_argument("--folder", default="INBOX", help="IMAPフォルダ（デフォルト: INBOX）")
    fetch_parser.add_argument("--limit", type=int, default=5, help="取得件数（最新から）")

    analyze_parser = sub.add_parser("analyze", help="過去n日間のメール本文を解析し、ポイント案件を抽出")
    analyze_parser.add_argument("--folder", default="INBOX", help="IMAPフォルダ（デフォルト: INBOX）")
    analyze_parser.add_argument("--days", type=int, default=7, help="さかのぼる日数（デフォルト: 7）")
    analyze_parser.add_argument("--output", help="抽出結果を保存するJSONファイルパス")
    analyze_parser.add_argument("--rejected-output", help="除外されたブロック一覧を保存するJSONファイルパス")

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

    if args.command == "analyze":
        days = max(args.days, 1)
        messages = fetch_messages_since(config, folder=args.folder, days=days)
        if not messages:
            print(f"{args.folder} には過去{days}日分のメールがありませんでした。")
            return

        campaigns, rejected_blocks = find_point_campaigns(messages)
        if campaigns:
            for idx, campaign in enumerate(campaigns, start=1):
                print("=" * 60)
                block_label = f"{campaign['id']}#{campaign['block_index'] + 1}"
                print(f"[{idx}] ID    : {block_label}")
                print(f"Date  : {campaign['date']}")
                print(f"From  : {campaign['from']}")
                if campaign.get("email_subject"):
                    print(f"Email : {campaign['email_subject']}")
                print(f"Subj  : {campaign['subject']}")
                print(f"Genre : {', '.join(campaign['genres'])}")
                print(f"Limit : {', '.join(campaign['matched_limit'])}")
                print(f"Point : {', '.join(campaign['matched_point'])}")
                print(f"Body  : {campaign['snippet']}")
        else:
            print(f"過去{days}日間では期間限定のポイント増案件は見つかりませんでした。")

        print("=" * 60)
        print(f"検出件数: {len(campaigns)} ブロック / メール {len(messages)} 通")
        print(f"除外ブロック: {len(rejected_blocks)}")

        if args.output:
            output_path = Path(args.output)
            output_path.write_text(json.dumps(campaigns, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"保存しました: {output_path}")

        if args.rejected_output:
            reject_path = Path(args.rejected_output)
            reject_path.write_text(json.dumps(rejected_blocks, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"除外リストを保存しました: {reject_path}")
        return

    if args.command == "send-test":
        to_addr = args.to or config.mail_address
        send_test_email(config, to_address=to_addr, subject=args.subject, body=args.body)
        print(f"SMTP送信完了: {config.mail_address} -> {to_addr}")
        return


if __name__ == "__main__":
    main()
