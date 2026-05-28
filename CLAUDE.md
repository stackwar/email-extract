# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

A Python script that connects to a QQ Mail IMAP inbox, finds emails with subject containing "聊天记录", extracts structured rider (骑手) records from the email body or inline images (via OCR), deduplicates by name+phone, and appends them to an Excel file.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run extraction once (process new emails then exit)
python extract.py

# Watch mode (IMAP IDLE, auto-process on new email)
python extract.py watch
```

## Architecture

Single-file script (`extract.py`) with two modes:
- **Normal mode**: connect, process unread matching emails, exit
- **Watch mode** (`watch` arg): IMAP IDLE long-poll, auto-reconnect, processes new emails as they arrive

Flow:
1. Connects to QQ Mail via IMAP SSL (credentials from `config.json`)
2. Searches inbox for emails with "聊天记录" in subject via UID search
3. Parses email body by date separator lines (`————— YYYY-MM-DD —————`) or Chinese date format (`YYYY年M月D日`)
4. Falls back to easyocr for inline images when text body yields no records
5. Deduplicates by name+phone, appends new records to `骑手信息.xlsx`
6. Tracks processed UIDs in `processed_ids.json` for incremental runs

## Key Details

- Uses `openpyxl` for Excel output, `easyocr` for image OCR
- Record delimiter regex: `(?:骑手)?姓名` to handle missing characters
- Date is per-record (from in-body date lines), not per-email
- Config in `config.json` (gitignored), see `config.example.json` for template
- CI/CD: tag push (`v*`) triggers GitHub Actions to build Windows EXE via PyInstaller
