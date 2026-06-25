"""Recipient validation, report matching, and email draft helpers."""

from __future__ import annotations

import base64
import html
import re
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path
from typing import Iterable
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
from independentsoft.msg import Attachment, Message, MessageFlag, Recipient, RecipientType

from config import (
    DEFAULT_PERIOD,
    DRAFT_SENDER_EMAIL,
    DRAFT_SENDER_NAME,
    HTML_TEMPLATE,
    MATCH_THRESHOLD,
    MAX_RECIPIENTS,
    SUBJECT_TEMPLATE,
    TEXT_TEMPLATE,
)


REQUIRED_COLUMNS = ["first_name", "email", "org_name"]
OUTPUT_COLUMNS = ["first_name", "last_name", "recipient_name", "email", "org_name", "contact_title"]
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
SIGNATURE_LOGO_PATH = Path(__file__).with_name("signature_logo.png")
SIGNATURE_WEBSITE_URL = "http://dallasfoundation.org/"
SIGNATURE_LINKEDIN_URL = "https://www.linkedin.com/in/erik-moss-a636125b/"

CONTACT_COLUMN_ALIASES = {
    "primary contact": "recipient_name",
    "primarycontact": "recipient_name",
    "contact": "recipient_name",
    "contact name": "recipient_name",
    "name": "recipient_name",
    "recipient name": "recipient_name",
    "greeting": "first_name",
    "first name": "first_name",
    "firstname": "first_name",
    "last name": "last_name",
    "lastname": "last_name",
    "contact title": "contact_title",
    "title": "contact_title",
    "contact email": "email",
    "email": "email",
    "email address": "email",
    "organization name": "org_name",
    "organization": "org_name",
    "org": "org_name",
    "org name": "org_name",
    "company": "org_name",
}


@dataclass(frozen=True)
class ReportFile:
    filename: str
    display_name: str
    normalized_name: str
    content: bytes


def format_attachment_filename(org_name: str, period: str) -> str:
    """Return a formatted PDF attachment filename for the given org and period."""
    org_clean = org_name.strip().title()
    safe_org = re.sub(r'[<>:"/\\|?*]', "", org_clean).strip()
    safe_period = re.sub(r'[<>:"/\\|?*]', "", period.strip())
    return f"{safe_org} Campus Engagement Report {safe_period}.pdf"


def normalize_org_name(value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    text = text.lower()
    text = re.sub(r"\b(attendance|engagement|campus|report|reports|pdf)\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def report_display_name(filename: str) -> str:
    name = Path(filename).stem
    name = re.sub(r"[_-]+", " ", name)
    name = re.sub(r"\b(attendance|engagement|campus|report|reports)\b", " ", name, flags=re.I)
    name = " ".join(name.split())
    return name.title()


def extract_reports_from_zip(uploaded_zip: bytes | BytesIO) -> list[ReportFile]:
    source = uploaded_zip if isinstance(uploaded_zip, BytesIO) else BytesIO(uploaded_zip)
    reports: list[ReportFile] = []
    with ZipFile(source) as report_zip:
        for name in report_zip.namelist():
            if not name.lower().endswith(".pdf"):
                continue
            if Path(name).name.lower().startswith("internal_summary"):
                continue
            content = report_zip.read(name)
            display = report_display_name(name)
            reports.append(
                ReportFile(
                    filename=name,
                    display_name=display,
                    normalized_name=normalize_org_name(display),
                    content=content,
                )
            )
    return reports


def _rapidfuzz_score(query: str, candidate: str) -> float | None:
    try:
        from rapidfuzz import fuzz
        return float(fuzz.token_set_ratio(query, candidate))
    except Exception:
        return None


def _difflib_score(query: str, candidate: str) -> float:
    from difflib import SequenceMatcher
    query_tokens = set(query.split())
    candidate_tokens = set(candidate.split())
    overlap = len(query_tokens & candidate_tokens) / max(len(query_tokens | candidate_tokens), 1)
    ratio = SequenceMatcher(None, query, candidate).ratio()
    return max(ratio, overlap) * 100


def fuzzy_score(query: str, candidate: str) -> float:
    rapid_score = _rapidfuzz_score(query, candidate)
    if rapid_score is not None:
        return rapid_score
    return _difflib_score(query, candidate)


def match_report(org_name: str, reports: Iterable[ReportFile], threshold: int = MATCH_THRESHOLD) -> dict:
    normalized = normalize_org_name(org_name)
    best_report: ReportFile | None = None
    best_score = 0.0
    for report in reports:
        score = fuzzy_score(normalized, report.normalized_name)
        if score > best_score:
            best_report = report
            best_score = score
    is_ready = best_report is not None and best_score >= threshold
    return {
        "matched_report": best_report.filename if is_ready and best_report else "",
        "matched_display_name": best_report.display_name if is_ready and best_report else "",
        "match_score": round(best_score, 1),
        "status": "Ready" if is_ready else "No match found",
        "report_content": best_report.content if is_ready and best_report else b"",
    }


def _canonical_column_name(column: object) -> str:
    text = "" if pd.isna(column) else str(column)
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def _split_name(full_name: str) -> tuple[str, str]:
    pieces = str(full_name or "").strip().split()
    if not pieces:
        return "", ""
    if len(pieces) == 1:
        return pieces[0], ""
    return pieces[0], " ".join(pieces[1:])


def normalize_recipient_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map either sample-recipient columns or the CEO list columns into app fields."""
    data = df.copy()
    rename_map = {}
    for column in data.columns:
        canonical = _canonical_column_name(column)
        if canonical in CONTACT_COLUMN_ALIASES:
            rename_map[column] = CONTACT_COLUMN_ALIASES[canonical]
    data = data.rename(columns=rename_map)

    for column in OUTPUT_COLUMNS:
        if column not in data.columns:
            data[column] = ""

    for column in OUTPUT_COLUMNS:
        data[column] = data[column].fillna("").astype(str).str.strip()

    if "recipient_name" in data.columns:
        missing_first = data["first_name"].eq("")
        data.loc[missing_first, "first_name"] = data.loc[missing_first, "recipient_name"].map(lambda value: _split_name(value)[0])
        missing_last = data["last_name"].eq("")
        data.loc[missing_last, "last_name"] = data.loc[missing_last, "recipient_name"].map(lambda value: _split_name(value)[1])

    data["recipient_name"] = data["recipient_name"].where(
        data["recipient_name"].ne(""),
        (data["first_name"] + " " + data["last_name"]).str.strip(),
    )

    return data[OUTPUT_COLUMNS].copy()


def validate_recipients(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    data = normalize_recipient_columns(df)
    errors: list[str] = []
    missing = [column for column in REQUIRED_COLUMNS if column not in data.columns]
    if missing:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), [f"Missing required columns: {', '.join(missing)}"]

    data = data[(data[REQUIRED_COLUMNS] != "").any(axis=1)].copy()

    if len(data) > MAX_RECIPIENTS:
        errors.append(f"Only the first {MAX_RECIPIENTS} recipients were kept.")
        data = data.head(MAX_RECIPIENTS).copy()

    for idx, row in data.iterrows():
        row_number = idx + 2
        if not row["first_name"]:
            errors.append(f"Row {row_number}: first_name is required.")
        if not row["email"] or not EMAIL_PATTERN.match(row["email"]):
            errors.append(f"Row {row_number}: email is missing or invalid.")
        if not row["org_name"]:
            errors.append(f"Row {row_number}: organization name is required.")

    return data.reset_index(drop=True), errors


def read_recipients_csv(uploaded_file) -> tuple[pd.DataFrame, list[str]]:
    name = getattr(uploaded_file, "name", "").lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file)
    return validate_recipients(df)


def build_match_table(recipients: pd.DataFrame, reports: list[ReportFile]) -> pd.DataFrame:
    rows = []
    for _, recipient in recipients.iterrows():
        match = match_report(recipient["org_name"], reports)
        rows.append(
            {
                "include": match["status"] == "Ready",
                "first_name": recipient["first_name"],
                "last_name": recipient["last_name"],
                "recipient_name": recipient["recipient_name"],
                "email": recipient["email"],
                "org_name": recipient["org_name"],
                "matched_report": match["matched_report"],
                "match_score": match["match_score"],
                "status": match["status"],
                "subject": render_subject(recipient["first_name"], recipient["org_name"], DEFAULT_PERIOD),
                "body": render_text_body(recipient["first_name"], recipient["org_name"], DEFAULT_PERIOD),
                "cc": "",
                "bcc": "",
            }
        )
    return pd.DataFrame(rows)


def render_subject(first_name: str, org_name: str, period: str) -> str:
    return SUBJECT_TEMPLATE.format(first_name=first_name, org_name=org_name, period=period)


def render_text_body(first_name: str, org_name: str, period: str) -> str:
    return TEXT_TEMPLATE.format(first_name=first_name, org_name=org_name, period=period)


def text_to_html(text: str) -> str:
    paragraphs = [
        html.escape(part.strip()).replace("\n", "<br>")
        for part in text.split("\n\n")
        if part.strip()
    ]
    return "\n".join(f"<p>{paragraph}</p>" for paragraph in paragraphs)


def _signature_logo_data_uri() -> str:
    if not SIGNATURE_LOGO_PATH.exists():
        return ""
    encoded = base64.b64encode(SIGNATURE_LOGO_PATH.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def render_signature_html() -> str:
    """Return the branded Outlook-style email signature shown below every draft."""
    logo_uri = _signature_logo_data_uri()
    logo_html = ""
    if logo_uri:
        logo_html = (
            f'<img src="{logo_uri}" alt="The Dallas Foundation" width="92" '
            'style="display:block;width:92px;height:auto;border:0;outline:none;text-decoration:none;">'
        )

    return f"""
<div style="margin-top:24px;font-family:Georgia,'Times New Roman',serif;color:#6f6f6f;line-height:1.35;">
  <div style="margin:0 0 28px 0;">{logo_html}</div>
  <div style="font-size:19px;font-weight:700;color:#4a4a4a;margin-bottom:3px;">Erik Moss</div>
  <div style="font-size:17px;color:#7a7a7a;margin-bottom:28px;">Director of the Water Cooler at Pegasus Park</div>
  <div style="font-size:14px;color:#777777;margin-bottom:10px;">
    3000 Pegasus Park Drive. #930&nbsp;&nbsp;|&nbsp;&nbsp;Dallas, TX 75247
  </div>
  <div style="font-size:14px;color:#777777;margin-bottom:10px;">
    <strong style="color:#626262;">P:</strong> 214-694-2529&nbsp;&nbsp;|&nbsp;&nbsp;<strong style="color:#626262;">C:</strong> 817-987-9945
  </div>
  <div style="font-size:14px;margin-bottom:8px;">
    <a href="{SIGNATURE_WEBSITE_URL}" style="color:#0000EE;text-decoration:underline;">dallasfoundation.org</a>
  </div>
  <div style="font-size:14px;">
    <a href="{SIGNATURE_LINKEDIN_URL}" style="color:#0000EE;text-decoration:underline;">{SIGNATURE_LINKEDIN_URL}</a>
  </div>
</div>
""".strip()


def render_signature_text() -> str:
    return f"""Erik Moss
Director of the Water Cooler at Pegasus Park

3000 Pegasus Park Drive. #930 | Dallas, TX 75247
P: 214-694-2529 | C: 817-987-9945
dallasfoundation.org
{SIGNATURE_LINKEDIN_URL}"""


def build_email_html(body: str) -> str:
    return f"{text_to_html(body)}\n{render_signature_html()}"


def build_email_text(body: str) -> str:
    return f"{body.rstrip()}\n\n{render_signature_text()}"


def _split_addresses(addresses: str) -> list[str]:
    if not addresses or not addresses.strip():
        return []
    return [a.strip() for a in addresses.split(",") if a.strip()]


def _set_optional_attr(target: object, name: str, value: object) -> None:
    try:
        setattr(target, name, value)
    except Exception:
        pass


def _format_display_address(display_name: str, email_address: str) -> str:
    name = str(display_name or "").strip()
    email = str(email_address or "").strip()
    if not name or name.lower() == email.lower():
        return email
    return f"{name} <{email}>"


def _make_recipient(display_name: str, email_address: str, recipient_type) -> Recipient:
    recipient = Recipient()
    formatted = _format_display_address(display_name, email_address)
    recipient.display_name = formatted
    recipient.email_address = email_address
    recipient.recipient_type = recipient_type
    _set_optional_attr(recipient, "smtp_address", email_address)
    _set_optional_attr(recipient, "address", email_address)
    _set_optional_attr(recipient, "address_type", "SMTP")
    return recipient


def build_msg_message(
    sender_email: str,
    sender_name: str,
    recipient_email: str,
    recipient_name: str,
    subject: str,
    text_body: str,
    html_body: str,
    attachment_filename: str,
    attachment_content: bytes,
    cc: str = "",
    bcc: str = "",
) -> Message:
    message = Message()
    message.message_class = "IPM.Note"
    message.message_flags = [MessageFlag.UNSENT]
    message.subject = subject
    message.body = text_body
    message.body_html_text = html_body
    message.sender_name = sender_name
    message.sender_email_address = sender_email
    _set_optional_attr(message, "sender_smtp_address", sender_email)

    recipients = []

    to_display = _format_display_address(recipient_name, recipient_email)
    recipients.append(_make_recipient(recipient_name, recipient_email, RecipientType.TO))

    for addr in _split_addresses(cc):
        recipients.append(_make_recipient(addr, addr, RecipientType.CC))

    for addr in _split_addresses(bcc):
        recipients.append(_make_recipient(addr, addr, RecipientType.BCC))

    message.recipients = recipients
    _set_optional_attr(message, "display_to", to_display)
    _set_optional_attr(message, "display_cc", "; ".join(_split_addresses(cc)))
    _set_optional_attr(message, "display_bcc", "; ".join(_split_addresses(bcc)))

    attachment = Attachment()
    attachment.file_name = attachment_filename
    attachment.data = attachment_content
    message.attachments = [attachment]

    return message


def _timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def build_log_row(row: dict, report_filename: str, status: str) -> dict:
    return {
        "Recipient": row.get("recipient_name", ""),
        "Org": row.get("org_name", ""),
        "Email": row.get("email", ""),
        "CC": row.get("cc", ""),
        "BCC": row.get("bcc", ""),
        "Report Attached": Path(report_filename).name if report_filename else "",
        "Status": status,
        "Timestamp": _timestamp(),
    }


def rows_to_generate(edited_rows: pd.DataFrame) -> pd.DataFrame:
    if edited_rows.empty:
        return edited_rows
    return edited_rows[
        (edited_rows["include"] == True) & (edited_rows["status"] == "Ready")
    ].copy()


def generate_eml_zip(
    edited_rows: pd.DataFrame,
    reports: list[ReportFile],
    period: str,
    sender_email: str = DRAFT_SENDER_EMAIL,
    sender_name: str = DRAFT_SENDER_NAME,
) -> tuple[bytes, pd.DataFrame]:
    report_lookup = {report.filename: report for report in reports}
    log_rows = []
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as zip_file:
        for _, row in rows_to_generate(edited_rows).iterrows():
            row_dict = row.to_dict()
            report = report_lookup.get(row["matched_report"])
            if report is None:
                log_rows.append(build_log_row(row_dict, "", "Skipped - missing report"))
                continue
            body = row.get("body") or render_text_body(row["first_name"], row["org_name"], period)
            text_body = build_email_text(body)
            html_body = build_email_html(body)
            attachment_name = format_attachment_filename(row["org_name"], period)
            message = build_msg_message(
                sender_email=sender_email,
                sender_name=sender_name,
                recipient_email=row["email"],
                recipient_name=row["recipient_name"],
                subject=row["subject"],
                text_body=text_body,
                html_body=html_body,
                attachment_filename=attachment_name,
                attachment_content=report.content,
                cc=row.get("cc", ""),
                bcc=row.get("bcc", ""),
            )
            org_slug = normalize_org_name(row["org_name"]).replace(" ", "_")
            name_slug = normalize_org_name(row["recipient_name"]).replace(" ", "_")
            filename = f"{org_slug}_{name_slug}.msg"
            zip_file.writestr(filename, message.to_bytes())
            log_rows.append(build_log_row(row_dict, report.filename, "Draft generated"))
    return output.getvalue(), pd.DataFrame(log_rows)


def log_to_csv(log_df: pd.DataFrame) -> bytes:
    output = StringIO()
    log_df.to_csv(output, index=False)
    return output.getvalue().encode("utf-8")


IMPORT_SCRIPT = """\
import glob, os, sys

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is not installed. Run:  pip install pywin32")
    sys.exit(1)

script_dir = os.path.dirname(os.path.abspath(__file__))
msg_files = sorted(glob.glob(os.path.join(script_dir, "*.msg")))

if not msg_files:
    print("No .msg files found in this folder.")
    sys.exit(0)

print(f"Found {len(msg_files)} draft(s). Connecting to Outlook...")

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    drafts = mapi.GetDefaultFolder(16)
except Exception as e:
    print(f"ERROR: Could not connect to Outlook. Make sure Outlook is open.\\n{e}")
    sys.exit(1)

success, failed = 0, 0
for path in msg_files:
    name = os.path.basename(path)
    try:
        item = outlook.CreateItemFromTemplate(path)
        item.Save()
        print(f"  done  {name}")
        success += 1
    except Exception as e:
        print(f"  FAIL  {name}  -  {e}")
        failed += 1

print(f"\\n{success} draft(s) saved to Outlook Drafts. {failed} failed.")
print("Open Outlook > Drafts to review and send.")
"""


def _generate_import_script() -> bytes:
    return IMPORT_SCRIPT.encode("utf-8")
