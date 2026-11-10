"""
Email Validator — pure logic.
No tkinter here.
"""

import re
import os
import pandas as pd
from openpyxl import Workbook


EMAIL_PATTERN = re.compile(
    r"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}"
    r"[a-zA-Z0-9])?"
    r"(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}"
    r"[a-zA-Z0-9])?)+$"
)


def is_valid(email):
    """Return True if the string strictly matches EMAIL_PATTERN."""
    return bool(email) and bool(EMAIL_PATTERN.match(str(email).strip()))


def validate_emails(emails):
    """Return a list of (email_str, valid_bool) tuples."""
    result = []
    for e in emails:
        s = "" if e is None else str(e).strip()
        result.append((s, is_valid(s)))
    return result


def split_pasted_text(text):
    """Split pasted text into email candidates."""
    if not text:
        return []
    tokens = re.split(r"[\s,;]+", text)
    out, seen = [], set()
    for t in tokens:
        t = t.strip().strip("<>\"'()[]{}.,;:")
        if "@" in t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def get_headers(file_path):
    """Return column headers of an Excel/CSV file."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        try:
            df = pd.read_csv(file_path, nrows=0, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, nrows=0, encoding="latin-1")
    else:
        df = pd.read_excel(file_path, nrows=0)
    return list(df.columns)


def detect_email_column(headers):
    """Return the first header that looks like an email column."""
    for h in headers:
        low = str(h).lower()
        if "email" in low or "mail" in low:
            return h
    return None


def load_emails(file_path, column):
    """Load emails from a column. Auto-detects the column if not given."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        try:
            df = pd.read_csv(file_path, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, encoding="latin-1")
    else:
        df = pd.read_excel(file_path)

    if not column or column not in df.columns:
        column = detect_email_column(list(df.columns))
    if not column:
        raise ValueError("No email column found in this file.")

    emails = df[column].fillna("").astype(str).tolist()
    return emails, column


def save_results(results, out_path):
    """results = list of (email, valid_bool). Writes an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.cell(row=1, column=1, value="Email")
    ws.cell(row=1, column=2, value="Status")
    for i, (email, valid) in enumerate(results, start=2):
        ws.cell(row=i, column=1, value=email)
        ws.cell(row=i, column=2, value="Valid" if valid else "Invalid")
    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 12
    wb.save(out_path)