"""Editable settings for the Water Cooler CEO email draft tool."""

from __future__ import annotations

MAX_RECIPIENTS = 50
MATCH_THRESHOLD = 82
DEFAULT_PERIOD = "January - June 2026"

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

DRAFT_SENDER_NAME = "Erik Moss"
DRAFT_SENDER_EMAIL = "erik@watercooler.org"

SUBJECT_TEMPLATE = "Your Pegasus Park Campus Engagement Report - {period}"

TEXT_TEMPLATE = """Hi {first_name},

In the spirit of transparency, we wanted to share something we've been tracking across our campus community.

As stewards of the Water Cooler - a space we're incredibly proud of - we're always looking at how it's being used and how we can continue to make it valuable for the organizations that call it home. Campus engagement is one of the key factors we look at as we assess the life and energy of this beautiful space.

Attached is {org_name}'s campus engagement report for {period}. It reflects your team's on-campus presence over the period and is part of our effort to keep every tenant informed and connected to how they're showing up here.

We'd love to hear your thoughts, and as always, please don't hesitate to reach out if you have any questions.

Warm regards,
Erik
"""

HTML_TEMPLATE = """\
<p>Hi {first_name},</p>

<p>In the spirit of transparency, we wanted to share something we've been tracking across our campus community.</p>

<p>As stewards of the Water Cooler - a space we're incredibly proud of - we're always looking at how it's being used and how we can continue to make it valuable for the organizations that call it home. Campus engagement is one of the key factors we look at as we assess the life and energy of this beautiful space.</p>

<p>Attached is {org_name}'s campus engagement report for {period}. It reflects your team's on-campus presence over the period and is part of our effort to keep every tenant informed and connected to how they're showing up here.</p>

<p>We'd love to hear your thoughts, and as always, please don't hesitate to reach out if you have any questions.</p>

<p>Warm regards,<br>
Erik</p>
"""
