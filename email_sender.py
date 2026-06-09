"""Recipient validation, report matching, email drafting, and delivery helpers."""

from __future__ import annotations

import os
import re
import smtplib
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from io import BytesIO, StringIO
from pathlib import Path
from typing import Iterable
from zipfile import ZipFile

import pandas as pd

from config import (
    DEFAULT_PERIOD,
    HTML_TEMPLATE,
    MATCH_THRESHOLD,
    MAX_RECIPIENTS,
    SMTP_ENV_VARS,
    SUBJECT_TEMPLATE,
    TEXT_TEMPLATE,
)


REQUIRED_COLUMNS = ["first_name", "last_name", "email", "org_name"]
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


@dataclass(frozen=True)
class ReportFile:
    filename: str
    display_name: str
    normalized_name: str
    content: bytes


@dataclass(frozen=True)
class SmtpSettings:
    sender_email: str
    sender_name: str
    smtp_host: str
    smtp_port: int
    smtp_password: str


def load_dotenv_file(path: str | Path = ".env") -> None:
    """Load .env values without requiring python-dotenv at import time."""
    try:
        from dotenv import load_dotenv

        load_dotenv(path)
        return
    except Exception:
        pass

    env_path = Path(path)
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def get_smtp_settings(
    path: str | Path = ".env",
    secrets: dict[str, object] | None = None,
) -> tuple[SmtpSettings | None, list[str]]:
    load_dotenv_file(path)
    secrets = secrets or {}
    values = {
        name: os.getenv(name) or str(secrets.get(name, "")).strip()
        for name in SMTP_ENV_VARS
    }
    missing = [name for name in SMTP_ENV_VARS if not values[name]]
    if missing:
        return None, missing
    try:
        port = int(values["SMTP_PORT"])
    except ValueError:
        return None, ["SMTP_PORT must be a number"]
    return (
        SmtpSettings(
            sender_email=values["SENDER_EMAIL"],
            sender_name=values["SENDER_NAME"],
            smtp_host=values["SMTP_HOST"],
            smtp_port=port,
            smtp_password=values["SMTP_PASSWORD"],
        ),
        [],
    )


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


def validate_recipients(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    data = df.copy()
    data.columns = [str(col).strip().lower() for col in data.columns]
    errors: list[str] = []
    missing = [column for column in REQUIRED_COLUMNS if column not in data.columns]
    if missing:
        return pd.DataFrame(columns=REQUIRED_COLUMNS), [f"Missing required columns: {', '.join(missing)}"]

    data = data[REQUIRED_COLUMNS].copy()
    for column in REQUIRED_COLUMNS:
        data[column] = data[column].fillna("").astype(str).str.strip()
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
            errors.append(f"Row {row_number}: org_name is required.")

    data["recipient_name"] = (data["first_name"] + " " + data["last_name"]).str.strip()
    return data.reset_index(drop=True), errors


def read_recipients_csv(uploaded_file) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_csv(uploaded_file)
    return validate_recipients(df)


def build_match_table(recipients: pd.DataFrame, reports: list[ReportFile]) -> pd.DataFrame:
    rows = []
    for _, recipient in recipients.iterrows():
        match = match_report(recipient["org_name"], reports)
        rows.append(
            {
                "send": match["status"] == "Ready",
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
            }
        )
    return pd.DataFrame(rows)


def render_subject(first_name: str, org_name: str, period: str) -> str:
    return SUBJECT_TEMPLATE.format(first_name=first_name, org_name=org_name, period=period)


def render_text_body(first_name: str, org_name: str, period: str) -> str:
    return TEXT_TEMPLATE.format(first_name=first_name, org_name=org_name, period=period)


def render_html_body(first_name: str, org_name: str, period: str) -> str:
    return HTML_TEMPLATE.format(first_name=first_name, org_name=org_name, period=period)


def text_to_html(text: str) -> str:
    paragraphs = [part.strip().replace("\n", "<br>") for part in text.split("\n\n") if part.strip()]
    return "\n".join(f"<p>{paragraph}</p>" for paragraph in paragraphs)


def build_email_message(
    sender_email: str,
    sender_name: str,
    recipient_email: str,
    recipient_name: str,
    subject: str,
    text_body: str,
    html_body: str,
    attachment_filename: str,
    attachment_content: bytes,
) -> EmailMessage:
    message = EmailMessage()
    message["From"] = formataddr((sender_name, sender_email))
    message["To"] = formataddr((recipient_name, recipient_email))
    message["Subject"] = subject
    message["Message-ID"] = make_msgid(domain=sender_email.split("@")[-1])
    message.set_content(text_body)
    message.add_alternative(html_body, subtype="html")
    message.add_attachment(
        attachment_content,
        maintype="application",
        subtype="pdf",
        filename=Path(attachment_filename).name,
    )
    return message


def _timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def build_log_row(row: dict, report_filename: str, status: str) -> dict:
    return {
        "Recipient": row.get("recipient_name", ""),
        "Org": row.get("org_name", ""),
        "Email": row.get("email", ""),
        "Report Attached": Path(report_filename).name if report_filename else "",
        "Status": status,
        "Timestamp": _timestamp(),
    }


def rows_to_send(edited_rows: pd.DataFrame) -> pd.DataFrame:
    if edited_rows.empty:
        return edited_rows
    return edited_rows[(edited_rows["send"] == True) & (edited_rows["status"] == "Ready")].copy()


def send_emails_via_smtp(
    edited_rows: pd.DataFrame,
    reports: list[ReportFile],
    settings: SmtpSettings,
    period: str,
) -> pd.DataFrame:
    report_lookup = {report.filename: report for report in reports}
    log_rows = []
    selected_rows = rows_to_send(edited_rows)

    with smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=30) as smtp:
        smtp.starttls()
        smtp.login(settings.sender_email, settings.smtp_password)
        for _, row in selected_rows.iterrows():
            row_dict = row.to_dict()
            report = report_lookup.get(row["matched_report"])
            if report is None:
                log_rows.append(build_log_row(row_dict, "", "Skipped - missing report"))
                continue
            try:
                body = row.get("body") or render_text_body(row["first_name"], row["org_name"], period)
                html = text_to_html(body)
                message = build_email_message(
                    sender_email=settings.sender_email,
                    sender_name=settings.sender_name,
                    recipient_email=row["email"],
                    recipient_name=row["recipient_name"],
                    subject=row["subject"],
                    text_body=body,
                    html_body=html,
                    attachment_filename=report.filename,
                    attachment_content=report.content,
                )
                smtp.send_message(message)
                log_rows.append(build_log_row(row_dict, report.filename, "Sent"))
            except Exception as exc:
                log_rows.append(build_log_row(row_dict, report.filename, f"Failed - {exc}"))
    return pd.DataFrame(log_rows)


def log_to_csv(log_df: pd.DataFrame) -> bytes:
    output = StringIO()
    log_df.to_csv(output, index=False)
    return output.getvalue().encode("utf-8")
