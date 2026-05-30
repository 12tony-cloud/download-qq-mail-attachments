#!/usr/bin/env python3
"""Download QQ Mail attachments and audit large-attachment links."""

from __future__ import annotations

import argparse
import getpass
import html
import imaplib
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass, asdict
from datetime import date, datetime, time as dt_time
from email import policy
from email.header import decode_header
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Iterable


IMAP_HOST = "imap.qq.com"
INVALID_CHARS = '<>:"/\\|?*\x00'
FTN_LINK_RE = re.compile(r"https://wx\.mail\.qq\.com/ftn/download\?[^\s\"'<>]+")
CLASS_FOLDER_NAMES = {
    "word": "Word\u6587\u6863",
    "zip": "zip\u538b\u7f29\u5305",
    "other": "\u5176\u4ed6\u9644\u4ef6",
}


@dataclass
class DownloadRecord:
    source: str
    uid: str
    subject: str
    name: str
    size: int
    dest: str
    status: str
    error: str = ""


def decode_mime_header(value: str | None) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for text, encoding in decode_header(value):
        if isinstance(text, bytes):
            parts.append(text.decode(encoding or "utf-8", errors="replace"))
        else:
            parts.append(text)
    return "".join(parts)


def safe_filename(name: str | None) -> str:
    value = decode_mime_header(name or "attachment").replace("\u00a0", " ").strip()
    for ch in INVALID_CHARS:
        value = value.replace(ch, "_")
    value = re.sub(r"\s+", " ", value).strip(" .")
    return value or "attachment"


def normalize_name(name: str) -> str:
    return re.sub(r"\s+", "", safe_filename(name)).casefold()


def target_folder(base_dir: Path, filename: str) -> Path:
    ext = Path(filename).suffix.lower()
    if ext in {".doc", ".docx"}:
        return base_dir / CLASS_FOLDER_NAMES["word"]
    if ext == ".zip":
        return base_dir / CLASS_FOLDER_NAMES["zip"]
    return base_dir / CLASS_FOLDER_NAMES["other"]


def existing_same_size(folder: Path, filename: str, size: int) -> Path | None:
    target = folder / filename
    if target.exists() and target.stat().st_size == size:
        return target
    return None


def unique_path(folder: Path, filename: str, size: int) -> tuple[Path, bool]:
    folder.mkdir(parents=True, exist_ok=True)
    existing = existing_same_size(folder, filename, size)
    if existing:
        return existing, True
    target = folder / filename
    if not target.exists():
        return target, False
    stem = target.stem
    suffix = target.suffix
    index = 1
    while True:
        candidate = folder / f"{stem} ({index}){suffix}"
        if not candidate.exists():
            return candidate, False
        if candidate.stat().st_size == size:
            return candidate, True
        index += 1


def parse_message_date(message) -> datetime | None:
    raw = message.get("Date")
    if not raw:
        return None
    try:
        parsed = parsedate_to_datetime(raw)
        if parsed.tzinfo:
            parsed = parsed.astimezone()
        return parsed.replace(tzinfo=None)
    except Exception:
        return None


def quote_mailbox(name: str) -> str:
    if name.upper() == "INBOX" or not re.search(r'\s|"|\\', name):
        return name
    return '"' + name.replace("\\", r"\\").replace('"', r'\"') + '"'


def mailbox_names(client: imaplib.IMAP4_SSL, inbox_only: bool) -> list[str]:
    if inbox_only:
        return ["INBOX"]
    status, boxes = client.list()
    if status != "OK":
        return ["INBOX"]
    names: list[str] = []
    for raw in boxes:
        line = raw.decode("utf-8", errors="replace")
        if "\\NoSelect" in line:
            continue
        match = re.search(r'\([^)]*\)\s+"[^"]*"\s+(".*"|[^ ]+)$', line)
        if not match:
            continue
        name = match.group(1)
        if name.startswith('"') and name.endswith('"'):
            name = name[1:-1].replace(r'\"', '"').replace(r"\\", "\\")
        names.append(name)
    return names or ["INBOX"]


def iter_attachment_parts(message) -> Iterable:
    for part in message.walk():
        if part.is_multipart():
            continue
        filename = part.get_filename()
        disposition = (part.get_content_disposition() or "").lower()
        if filename or disposition == "attachment":
            yield part


def iter_text_bodies(message) -> Iterable[str]:
    for part in message.walk() if message.is_multipart() else [message]:
        if part.is_multipart():
            continue
        if part.get_content_type() not in {"text/plain", "text/html"}:
            continue
        try:
            text = part.get_content()
        except Exception:
            text = (part.get_payload(decode=True) or b"").decode("utf-8", errors="replace")
        yield html.unescape(text)


def extract_ftn_links(message) -> list[str]:
    links: list[str] = []
    seen: set[tuple[str, str]] = set()
    for body in iter_text_bodies(message):
        for match in FTN_LINK_RE.finditer(body):
            url = match.group(0).rstrip(").,;")
            query = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
            key = (query.get("key") or query.get("k") or [""])[0]
            code = (query.get("code") or [""])[0]
            ident = (key, code)
            if key and code and ident not in seen:
                seen.add(ident)
                links.append(url)
    return links


def request_headers(referer: str, accept: str) -> dict[str, str]:
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/124 Safari/537.36"
        ),
        "Accept": accept,
        "Referer": referer,
    }


def resolve_large_attachment(url: str) -> dict:
    query = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
    key = (query.get("key") or query.get("k") or [""])[0]
    code = (query.get("code") or [""])[0]
    api_url = "https://wx.mail.qq.com/ftn/download?" + urllib.parse.urlencode(
        {"func": "3", "key": key, "code": code, "f": "json"}
    )
    headers = request_headers(url, "application/json, text/plain, */*")
    headers["X-Requested-With"] = "XMLHttpRequest"
    request = urllib.request.Request(api_url, headers=headers)
    with urllib.request.urlopen(request, timeout=60) as response:
        data = json.loads(response.read().decode("utf-8", errors="replace"))
    body = data.get("body") or {}
    ret = (data.get("head") or {}).get("ret")
    if ret != 0 or not body.get("url") or not body.get("name"):
        raise RuntimeError(f"QQ large-attachment API failed ret={ret}")
    return body


def download_exact(url: str, dest: Path, referer: str, expected_size: int) -> int:
    temp = dest.with_name(dest.name + ".part")
    if temp.exists():
        temp.unlink()
    headers = request_headers(
        referer,
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "application/octet-stream,*/*;q=0.8",
    )
    headers["Upgrade-Insecure-Requests"] = "1"
    request = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(request, timeout=120) as response, temp.open("wb") as handle:
        total = 0
        if expected_size <= 0:
            expected_size = int(response.headers.get("content-length") or 0)
        while expected_size <= 0 or total < expected_size:
            size = 512 * 1024 if expected_size <= 0 else min(512 * 1024, expected_size - total)
            chunk = response.read(size)
            if not chunk:
                break
            handle.write(chunk)
            total += len(chunk)
            if expected_size <= 0 and len(chunk) == 0:
                break
    sample = temp.read_bytes()[:80] if temp.exists() else b""
    if sample.lstrip().startswith((b"<!DOCTYPE", b"<html")):
        temp.unlink(missing_ok=True)
        raise RuntimeError(f"download returned HTML shell ({len(sample)} byte sample)")
    if expected_size > 0 and total != expected_size:
        temp.unlink(missing_ok=True)
        raise RuntimeError(f"size mismatch: got {total}, expected {expected_size}")
    temp.replace(dest)
    return total


def save_payload(base_dir: Path, filename: str, payload: bytes) -> tuple[Path, str]:
    safe = safe_filename(filename)
    folder = target_folder(base_dir, safe)
    target, exists = unique_path(folder, safe, len(payload))
    if exists:
        return target, "skipped_existing"
    target.write_bytes(payload)
    return target, "downloaded"


def collect_messages(account: str, auth_code: str, since: date, inbox_only: bool):
    since_start = datetime.combine(since, dt_time.min)
    imap_since = since.strftime("%d-%b-%Y")
    with imaplib.IMAP4_SSL(IMAP_HOST, 993) as client:
        client.login(account, auth_code)
        for box in mailbox_names(client, inbox_only):
            try:
                status, _ = client.select(quote_mailbox(box), readonly=True)
            except imaplib.IMAP4.error:
                continue
            if status != "OK":
                continue
            status, data = client.uid("search", None, "SINCE", imap_since)
            if status != "OK" or not data or not data[0]:
                continue
            for uid_bytes in data[0].split():
                uid = uid_bytes.decode("ascii", errors="replace")
                status, fetched = client.uid("fetch", uid_bytes, "(RFC822)")
                if status != "OK":
                    continue
                raw_message = next(
                    (item[1] for item in fetched if isinstance(item, tuple) and item[1]),
                    None,
                )
                if not raw_message:
                    continue
                message = BytesParser(policy=policy.default).parsebytes(raw_message)
                msg_date = parse_message_date(message)
                if msg_date and msg_date < since_start:
                    continue
                yield uid, decode_mime_header(message.get("Subject", "")), message


def download_all(account: str, auth_code: str, base_dir: Path, since: date, inbox_only: bool):
    records: list[DownloadRecord] = []
    errors: list[str] = []
    seen_normal: set[tuple[str, int]] = set()
    seen_large: set[tuple[str, str]] = set()
    expected_large: list[dict] = []

    for uid, subject, message in collect_messages(account, auth_code, since, inbox_only):
        for part in iter_attachment_parts(message):
            filename = safe_filename(part.get_filename() or "attachment")
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            fingerprint = (filename, len(payload))
            if fingerprint in seen_normal:
                continue
            seen_normal.add(fingerprint)
            try:
                target, status = save_payload(base_dir, filename, payload)
                records.append(
                    DownloadRecord("normal", uid, subject, filename, len(payload), str(target), status)
                )
            except Exception as exc:
                records.append(
                    DownloadRecord("normal", uid, subject, filename, len(payload), "", "error", str(exc))
                )

        for link in extract_ftn_links(message):
            query = urllib.parse.parse_qs(urllib.parse.urlparse(link).query)
            key = (query.get("key") or query.get("k") or [""])[0]
            code = (query.get("code") or [""])[0]
            ident = (key, code)
            if ident in seen_large:
                continue
            seen_large.add(ident)
            try:
                body = resolve_large_attachment(link)
                filename = safe_filename(body.get("name"))
                size = int(body.get("size") or 0)
                expected_large.append({"name": filename, "size": size, "uid": uid, "subject": subject})
                folder = target_folder(base_dir, filename)
                target, exists = unique_path(folder, filename, size)
                if exists:
                    status = "skipped_existing"
                else:
                    for attempt in range(1, 4):
                        try:
                            if attempt > 1:
                                body = resolve_large_attachment(link)
                            download_exact(body["url"], target, link, size)
                            break
                        except Exception:
                            if attempt == 3:
                                raise
                            time.sleep(1.5 * attempt)
                    status = "downloaded"
                records.append(DownloadRecord("large", uid, subject, filename, size, str(target), status))
            except Exception as exc:
                errors.append(f"uid={uid} subject={subject}: {exc}")
                records.append(DownloadRecord("large", uid, subject, "", 0, "", "error", str(exc)))

    return records, expected_large, errors


def local_files(base_dir: Path) -> list[Path]:
    return [p for p in base_dir.rglob("*") if p.is_file()]


def audit(base_dir: Path, expected_large: list[dict]) -> dict:
    files = local_files(base_dir)
    local_index = [
        {"name": p.name, "norm": normalize_name(p.name), "size": p.stat().st_size, "path": str(p)}
        for p in files
    ]
    missing_large = []
    for item in expected_large:
        expected_norm = normalize_name(item["name"])
        matches = [
            entry
            for entry in local_index
            if entry["size"] == item["size"] and entry["norm"] == expected_norm
        ]
        if not matches:
            missing_large.append(item)

    zip_docx_errors = []
    for path in files:
        if path.suffix.lower() not in {".zip", ".docx"}:
            continue
        try:
            with zipfile.ZipFile(path) as archive:
                bad = archive.testzip()
            if bad:
                zip_docx_errors.append({"path": str(path), "error": f"bad member {bad}"})
        except Exception as exc:
            zip_docx_errors.append({"path": str(path), "error": str(exc)})

    part_files = [str(path) for path in files if path.name.endswith(".part")]
    return {
        "missing_big_attachments": missing_large,
        "zip_docx_errors": zip_docx_errors,
        "part_files": part_files,
    }


def write_lists(base_dir: Path) -> None:
    configs = [
        (base_dir / CLASS_FOLDER_NAMES["word"], "Word\u6587\u6863\u540d\u5355.txt", "Word\u6587\u6863\u540d\u5355"),
        (base_dir / CLASS_FOLDER_NAMES["zip"], "zip\u538b\u7f29\u5305\u540d\u5355.txt", "zip\u538b\u7f29\u5305\u540d\u5355"),
    ]
    for folder, list_name, title in configs:
        folder.mkdir(parents=True, exist_ok=True)
        list_path = folder / list_name
        files = sorted([p for p in folder.iterdir() if p.is_file() and p.name != list_name], key=lambda p: p.name)
        lines = [title, "count: " + str(len(files)), ""]
        lines.extend(f"{index}. {path.name}" for index, path in enumerate(files, 1))
        list_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Download QQ Mail attachments and audit results.")
    parser.add_argument("--account", default=os.environ.get("QQMAIL_ACCOUNT"))
    parser.add_argument("--auth-code", default=os.environ.get("QQMAIL_AUTH_CODE"))
    parser.add_argument("--base-dir", default=str(Path.home() / "Desktop" / "\u4f5c\u4e1a"))
    parser.add_argument("--since", required=True, help="Include messages on or after YYYY-MM-DD.")
    parser.add_argument("--inbox-only", action="store_true", help="Scan only INBOX instead of all selectable folders.")
    parser.add_argument("--audit-only", action="store_true", help="Only audit existing files from the last report.")
    parser.add_argument("--write-lists", action="store_true", help="Write Word and zip filename lists.")
    args = parser.parse_args()

    base_dir = Path(args.base_dir)
    since = date.fromisoformat(args.since)
    report_path = base_dir / "attachment_download_report.json"

    if args.audit_only:
        if not report_path.exists():
            print(f"No report found: {report_path}", file=sys.stderr)
            return 2
        report = json.loads(report_path.read_text(encoding="utf-8"))
        expected_large = report.get("expected_large_attachments", [])
        records = [DownloadRecord(**item) for item in report.get("downloads", [])]
        errors = []
    else:
        account = args.account or input("QQ Mail account: ").strip()
        auth_code = args.auth_code or getpass.getpass("QQ Mail IMAP authorization code: ").strip()
        records, expected_large, errors = download_all(
            account=account,
            auth_code=auth_code,
            base_dir=base_dir,
            since=since,
            inbox_only=args.inbox_only,
        )

    audit_result = audit(base_dir, expected_large)
    if args.write_lists:
        write_lists(base_dir)

    downloads = [asdict(record) for record in records]
    report = {
        "since": since.isoformat(),
        "base_dir": str(base_dir),
        "downloads": downloads,
        "expected_large_attachments": expected_large,
        "errors": errors,
        "audit": audit_result,
        "counts": {
            "downloaded": sum(1 for record in records if record.status == "downloaded"),
            "skipped_existing": sum(1 for record in records if record.status == "skipped_existing"),
            "error": sum(1 for record in records if record.status == "error"),
        },
    }
    base_dir.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(report["counts"], ensure_ascii=False))
    print(json.dumps(audit_result, ensure_ascii=False))
    print(f"Report: {report_path}")
    return 1 if errors or audit_result["missing_big_attachments"] or audit_result["zip_docx_errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
