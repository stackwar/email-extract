# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

A Python script that connects to a QQ Mail IMAP inbox, finds the most recent email with subject containing "聊天记录", extracts structured rider (骑手) records from the email body, and writes them to an Excel file.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the extraction
python extract.py
```

## Architecture

Single-file script (`extract.py`) with this flow:
1. Connects to QQ Mail via IMAP SSL
2. Searches inbox for emails with "聊天记录" in subject (newest first)
3. Parses the email body using regex to extract rider records (name, phone, station, housing, job type, notes)
4. Writes all records to `骑手信息.xlsx` with the email date as a column

## Key Details

- Uses `openpyxl` for Excel output
- Email body is split into blocks by "骑手姓名" delimiter, then each block is parsed for known fields
- IMAP credentials are hardcoded in the script (QQ Mail auth code)
- Output file is `骑手信息.xlsx` in the project root
