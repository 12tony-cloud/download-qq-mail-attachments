---
name: download-qq-mail-attachments
description: Download and audit QQ Mail attachments with IMAP plus QQ large-attachment link handling. Use when Codex needs to fetch QQ Mail or QQ邮箱 attachments after a date, including normal MIME attachments and QQ "超大附件"/file-transfer-station links, then classify Word documents, zip archives, and other files into separate folders and verify nothing is missing.
---

# Download QQ Mail Attachments

## Workflow

1. Confirm the date range, destination folder, and classification rules.
2. Get explicit user consent before accessing email. Never persist QQ passwords, authorization codes, session IDs, `key/code` link tokens, or full large-attachment URLs in skill files, reports, or final answers.
3. Prefer `scripts/qq_mail_attachment_downloader.py` for the end-to-end workflow:

```powershell
$env:QQMAIL_ACCOUNT = "user@qq.com"
$env:QQMAIL_AUTH_CODE = Read-Host "QQ Mail IMAP authorization code"
python "C:\Users\23734\.codex\skills\download-qq-mail-attachments\scripts\qq_mail_attachment_downloader.py" `
  --since 2026-05-24 `
  --base-dir "C:\Users\23734\Desktop\作业" `
  --write-lists
```

Use `--account` or `QQMAIL_ACCOUNT`; use `--auth-code`, `QQMAIL_AUTH_CODE`, or the hidden prompt. Avoid putting the authorization code directly in files.

## What The Script Does

- Connects to `imap.qq.com` with IMAP SSL and scans messages on or after `--since`.
- Downloads normal MIME attachments.
- Extracts QQ large-attachment links from plain-text and HTML mail bodies.
- Resolves large attachments by calling the QQ endpoint with `func=3`, `key`, `code`, and `f=json`.
- Downloads `body.url` using the exact byte count returned by QQ instead of waiting for EOF, which avoids hangs on some file-transfer responses.
- Sorts files into:
  - `Word文档` for `.doc` and `.docx`
  - `zip压缩包` for `.zip`
  - `其他附件` for everything else
- Skips existing same-name, same-size files and creates numbered names for same-name, different-size files.
- Writes `attachment_download_report.json` in the base folder.
- With `--write-lists`, writes `Word文档名单.txt` and `zip压缩包名单.txt` into their matching folders.

## Validation Checklist

After running, inspect the script summary and report:

- `errors` should be empty.
- `missing_big_attachments` should be empty.
- `zip_docx_errors` should be empty for `.zip` and `.docx` files.
- `part_files` should be `0`.
- Re-run with the same arguments if uncertain; existing same-size files should be skipped.

If the user says attachments are still missing, run the script again with the same date and `--audit-only` to compare all large-attachment links in email bodies against local files by normalized name and size.

## Troubleshooting

- HTML page saved instead of the file: the large-attachment link was fetched directly. Resolve it first with `f=json`, then download the returned `body.url`.
- Download hangs: read exactly the expected `body.size` bytes and close the response once that count is reached.
- Chinese path is mangled in a shell pipeline: pass paths as command arguments, or use Python `Path` strings/Unicode escapes inside scripts.
- Folder scan misses mail: run without `--inbox-only` so all selectable IMAP folders are scanned.
- Duplicate-looking files: compare by exact byte size first; hash only when deciding whether to remove duplicates.
