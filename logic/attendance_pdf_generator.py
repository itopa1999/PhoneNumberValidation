import os
import re
import pandas as pd

# ---------- reportlab ----------
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
)
from reportlab.lib.enums import TA_CENTER


# ---------- Constants ----------
WEEKS = [f"Week{i:02d}" for i in range(1, 27)]
REQUIRED_SHEET = "Attendance_Report"
REQUIRED_COLUMNS = ["Name", "Phone", "Type", "Source"]


class AttendanceReportError(Exception):
    """Raised for any user-facing validation or build problem."""
    pass


# ---------- Helpers ----------
def sanitize_source(source: str) -> str:
    """Strip HTML, illegal filename chars, and collapse whitespace."""
    s = str(source)
    s = s.replace("<br>", " ").replace("<br/>", " ").replace("<BR>", " ")
    s = s.replace("<", " ").replace(">", " ")
    s = re.sub(r'[\\/:*?"|]', "_", s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def validate_week(week_col: str) -> None:
    if week_col not in WEEKS:
        raise AttendanceReportError(f"Invalid week column: {week_col!r}")


def week_to_module_name(week_col: str) -> str:
    """'Week03' → 'Module 03' — used in titles and headers."""
    return week_col.replace("Week", "Module ")


# ---------- Data loading ----------
def load_attendance_data(file_path: str, week_col: str) -> pd.DataFrame:
    """
    Load the 'Attendance_Report' sheet and validate required columns.
    Returns a cleaned DataFrame (all required cols converted to str).
    Raises AttendanceReportError on any validation problem.
    """
    validate_week(week_col)

    if not file_path or not os.path.exists(file_path):
        raise AttendanceReportError("Please select a valid Excel file.")

    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        raise AttendanceReportError(f"Could not open Excel file:\n{e}")

    if REQUIRED_SHEET not in xls.sheet_names:
        raise AttendanceReportError(
            "The Excel file must contain a sheet named "
            f"'{REQUIRED_SHEET}'.\n"
            f"Found sheets: {', '.join(xls.sheet_names)}"
        )

    try:
        df = pd.read_excel(file_path, sheet_name=REQUIRED_SHEET)
    except Exception as e:
        raise AttendanceReportError(f"Could not read sheet:\n{e}")

    required = REQUIRED_COLUMNS + [week_col]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise AttendanceReportError(
            f"Missing required columns: {', '.join(missing)}\n"
            f"(Is the week column '{week_col}' present in the sheet?)"
        )

    for c in required:
        df[c] = df[c].fillna("").astype(str)

    return df


# ---------- Grouping ----------
def group_by_source(df: pd.DataFrame, week_col: str):
    """
    Group by 'Source' (order preserved), sort each group Present → Absent →
    Other, drop groups whose source is empty/whitespace.
    Returns a list of (source_str, sub_dataframe).
    """
    def sort_key(val):
        s = str(val).strip().lower()
        if s == "present":
            return 0
        if s == "absent":
            return 1
        return 2

    groups = []
    for source, sub in df.groupby("Source", sort=False):
        sub = sub.copy()
        sub["__sort_key"] = sub[week_col].apply(sort_key)
        sub = (sub.sort_values(by="__sort_key")
                  .drop(columns=["__sort_key"])
                  .reset_index(drop=True))
        groups.append((source, sub))

    return [(s, g) for s, g in groups if str(s).strip()]


# ---------- Filename ----------
def generate_pdf_filename(source: str, week_col: str) -> str:
    safe = sanitize_source(source)
    module_name = week_col.replace("Week", "Module_")
    return f"generated_report_for_{safe}_{module_name}.pdf"


# ---------- PDF build ----------
def build_attendance_pdf(source: str, df: pd.DataFrame,
                         week_col: str, out_path: str) -> str:
    """
    Build a single PDF report for one Source group.
    Returns the written path. Raises AttendanceReportError on failure.
    """
    validate_week(week_col)
    module_name = week_to_module_name(week_col)

    try:
        doc = SimpleDocTemplate(
            out_path,
            pagesize=landscape(A4),
            leftMargin=0.5 * inch,
            rightMargin=0.5 * inch,
            topMargin=0.5 * inch,
            bottomMargin=0.5 * inch,
            title=f"Attendance Report - {module_name}",
        )
    except Exception as e:
        raise AttendanceReportError(f"Could not create PDF document:\n{e}")

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title", parent=styles["Title"],
        fontSize=15, alignment=TA_CENTER,
        textColor=colors.HexColor("#1a2b3c"),
    )
    subtitle_style = ParagraphStyle(
        "sub", parent=styles["Normal"],
        fontSize=11, alignment=TA_CENTER,
        textColor=colors.HexColor("#2a4b6c"),
    )

    elements = []
    elements.append(Paragraph(f"Attendance Report — {module_name}", title_style))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(sanitize_source(source), subtitle_style))
    elements.append(Spacer(1, 4))
    elements.append(Paragraph(f"Total candidates: {len(df)}", subtitle_style))
    elements.append(Spacer(1, 12))

    data = [["#", "Name", "Phone", "Type", module_name]]
    for i, (_, row) in enumerate(df.iterrows(), 1):
        data.append([
            str(i),
            str(row.get("Name", "") or ""),
            str(row.get("Phone", "") or ""),
            str(row.get("Type", "") or ""),
            str(row.get(week_col, "") or ""),
        ])

    table = Table(data, repeatRows=1, colWidths=[
        0.4 * inch, 3.6 * inch, 1.5 * inch, 1.2 * inch, 1.3 * inch
    ])

    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2a7de1")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#b0c0d0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#f0f4f8")]),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("ALIGN", (2, 1), (4, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]

    # Color the week/module cell per row: Present → green, Absent → red
    for i, (_, row) in enumerate(df.iterrows(), 1):
        status = str(row.get(week_col, "")).strip().lower()
        if status == "present":
            table_style.append(("BACKGROUND", (4, i), (4, i),
                                colors.HexColor("#d4edda")))
            table_style.append(("TEXTCOLOR", (4, i), (4, i),
                                colors.HexColor("#155724")))
        elif status == "absent":
            table_style.append(("BACKGROUND", (4, i), (4, i),
                                colors.HexColor("#f8d7da")))
            table_style.append(("TEXTCOLOR", (4, i), (4, i),
                                colors.HexColor("#721c24")))

    table.setStyle(TableStyle(table_style))
    elements.append(table)

    try:
        doc.build(elements)
    except Exception as e:
        raise AttendanceReportError(f"Could not build PDF:\n{e}")

    return out_path