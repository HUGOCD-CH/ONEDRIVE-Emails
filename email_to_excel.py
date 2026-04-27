#!/usr/bin/env python3
"""
Email to Excel Converter
Reads .txt email files (HTML bodies exported from Outlook) from a folder
and writes an Excel summary with Subject, Date, Body Excerpt, and a
Claude-generated reply suggestion.

Usage:
    pip install -r requirements.txt
    set ANTHROPIC_API_KEY=sk-...   (Windows) or export ANTHROPIC_API_KEY=sk-...
    python email_to_excel.py
"""

import os
import re
import sys
from datetime import datetime
from os.path import getmtime
from pathlib import Path

try:
    from bs4 import BeautifulSoup
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    import anthropic
except ImportError as exc:
    print(f"Missing dependency: {exc}")
    print("Run:  pip install -r requirements.txt")
    sys.exit(1)


# ── Configuration ──────────────────────────────────────────────────────────────
EMAILS_FOLDER = r"C:\Users\diash2\OneDrive - Medtronic PLC\TEMP\Emails"
BODY_EXCERPT_LENGTH = 300
OUTPUT_PREFIX = "Email_Summary"
AI_MODEL = "claude-haiku-4-5-20251001"

_SYSTEM_PROMPT = (
    "You are a professional email assistant. "
    "Given an email body, write a concise, polite reply suggestion in 2-3 sentences. "
    "Reply only with the suggested text — no preamble, no subject line, no sign-off."
)


# ── HTML → plain text ──────────────────────────────────────────────────────────
def html_to_text(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["style", "script", "meta", "head", "img"]):
        tag.decompose()
    text = soup.get_text(separator=" ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def make_excerpt(text: str) -> str:
    flat = " ".join(text.split())
    if len(flat) <= BODY_EXCERPT_LENGTH:
        return flat
    return flat[:BODY_EXCERPT_LENGTH].rstrip() + "…"


# ── AI reply suggestion ────────────────────────────────────────────────────────
def get_reply_suggestion(client: anthropic.Anthropic, plain_text: str) -> str:
    try:
        response = client.messages.create(
            model=AI_MODEL,
            max_tokens=256,
            system=[
                {
                    "type": "text",
                    "text": _SYSTEM_PROMPT,
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            messages=[{"role": "user", "content": plain_text[:4000]}],
        )
        return response.content[0].text.strip()
    except Exception as exc:
        return f"[Error: {exc}]"


# ── Parse one email file ───────────────────────────────────────────────────────
def parse_email_file(path: Path, client: anthropic.Anthropic) -> dict:
    content = path.read_text(encoding="utf-8", errors="replace")
    plain = html_to_text(content)
    file_date = datetime.fromtimestamp(getmtime(path)).strftime("%Y-%m-%d %H:%M")
    return {
        "Subject": path.stem,
        "Date": file_date,
        "Body Excerpt": make_excerpt(plain),
        "Reply Suggestion": get_reply_suggestion(client, plain),
    }


# ── Excel generation ───────────────────────────────────────────────────────────
COLUMNS = [
    ("Subject",          50),
    ("Date",             18),
    ("Body Excerpt",     60),
    ("Reply Suggestion", 60),
]

_HDR_FILL = PatternFill(start_color="00285A", end_color="00285A", fill_type="solid")
_HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
_ROW_FONT = Font(name="Calibri", size=10)
_ALT_FILL = PatternFill(start_color="EEF2F7", end_color="EEF2F7", fill_type="solid")
_WRAP     = Alignment(wrap_text=True, vertical="top")
_CENTER   = Alignment(horizontal="center", vertical="center", wrap_text=True)
_BORDER   = Border(
    bottom=Side(style="thin", color="CCCCCC"),
    right =Side(style="thin", color="CCCCCC"),
)


def generate_excel(emails: list, output_path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Email Summary"
    ws.freeze_panes = "A2"

    for col_idx, (header, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = _HDR_FONT
        cell.fill      = _HDR_FILL
        cell.alignment = _CENTER
        ws.column_dimensions[cell.column_letter].width = width
    ws.row_dimensions[1].height = 22

    for row_idx, email in enumerate(emails, 2):
        fill = _ALT_FILL if row_idx % 2 == 0 else None
        for col_idx, (header, _) in enumerate(COLUMNS, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=email.get(header, ""))
            cell.font      = _ROW_FONT
            cell.alignment = _WRAP
            cell.border    = _BORDER
            if fill:
                cell.fill = fill

    ws.auto_filter.ref = f"A1:{ws.cell(1, len(COLUMNS)).coordinate}"
    wb.save(output_path)


# ── Main ───────────────────────────────────────────────────────────────────────
def main() -> None:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: ANTHROPIC_API_KEY environment variable is not set.")
        print("Set it with:  set ANTHROPIC_API_KEY=sk-...")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    folder = Path(EMAILS_FOLDER)
    if not folder.exists():
        print(f"ERROR: Folder not found:\n  {folder}")
        sys.exit(1)

    txt_files = sorted(
        f for f in folder.glob("*.txt")
        if not f.name.startswith(OUTPUT_PREFIX)
    )

    if not txt_files:
        print(f"No .txt email files found in:\n  {folder}")
        sys.exit(0)

    print(f"Found {len(txt_files)} email file(s).\n")

    emails = []
    for path in txt_files:
        try:
            emails.append(parse_email_file(path, client))
            print(f"  OK  {path.name}")
        except Exception as exc:
            print(f"  ERR {path.name}: {exc}")

    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = folder / f"{OUTPUT_PREFIX}_{timestamp}.xlsx"
    generate_excel(emails, output_path)

    print(f"\nSaved: {output_path}")


if __name__ == "__main__":
    main()
