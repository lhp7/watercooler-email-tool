"""Editable settings for the Water Cooler CEO email tool."""

from __future__ import annotations

MAX_RECIPIENTS = 30
MATCH_THRESHOLD = 82
DEFAULT_PERIOD = "January - May 2026"

BRAND_COLORS = {
    "navy": "#153E5C",
    "blue": "#1F77B4",
    "teal": "#00A7A5",
    "aqua": "#BFEFED",
    "gold": "#F2B84B",
    "coral": "#E96B56",
    "ink": "#20323F",
    "muted": "#6B7A86",
    "paper": "#F7FBFC",
    "white": "#FFFFFF",
    "line": "#D9E7EC",
}

SENDER_FALLBACK_NAME = "Erik"
SENDER_FALLBACK_EMAIL = "erik@watercooler.org"

SMTP_ENV_VARS = [
    "SENDER_EMAIL",
    "SENDER_NAME",
    "SMTP_HOST",
    "SMTP_PORT",
    "SMTP_PASSWORD",
]

SUBJECT_TEMPLATE = "Your Pegasus Park Campus Engagement Report - {period}"

TEXT_TEMPLATE = """Hi {first_name},

I hope you're doing well! I'm reaching out from the Water Cooler team at Pegasus Park to share {org_name}'s campus engagement report for {period}.

Your report is attached and includes a summary of your team's on-campus attendance, badge activity, and engagement trends for the period. These reports are generated from our keycard access data and are sent every six months to help your team stay informed about how you're using the space.

Please don't hesitate to reach out if you have any questions about the data or would like to discuss your team's campus presence.

Warm regards,
Erik
Water Cooler at Pegasus Park | Pegasus Park
"""

HTML_TEMPLATE = """\
<p>Hi {first_name},</p>

<p>I hope you're doing well! I'm reaching out from the Water Cooler team at Pegasus Park to share {org_name}'s campus engagement report for {period}.</p>

<p>Your report is attached and includes a summary of your team's on-campus attendance, badge activity, and engagement trends for the period. These reports are generated from our keycard access data and are sent every six months to help your team stay informed about how you're using the space.</p>

<p>Please don't hesitate to reach out if you have any questions about the data or would like to discuss your team's campus presence.</p>

<p>Warm regards,<br>
Erik<br>
Water Cooler at Pegasus Park | Pegasus Park</p>
"""

