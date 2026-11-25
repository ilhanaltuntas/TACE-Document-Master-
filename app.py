"""TKPR2 Document Master List Dashboard — single-file version.

Run with current Python 3.15-compatible packages:
    python -m streamlit run app.py

Required packages are installed by run_dashboard.bat from requirements.txt.

The data source is the matching Excel file in the same folder as app.py.
The A1:P391 range in the Phase-2 Deliverables sheet is used.
Activity ID, Activity Name, Start, and Finish in columns F:I are excluded.
Plan Issue date is read directly from column M.
"""

from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
import html
import json
from pathlib import Path
from typing import BinaryIO

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SHEET_NAME = "Phase-2 Deliverables"
SOURCE_FILENAME = "Document Master List - TKPR2 (for RET review and update)_20260925 1.xlsx"
HEADER_ROW_COUNT = 3
SOURCE_MAX_ROWS = 391
READ_USECOLS = "A:P"
SOURCE_COLUMN_LETTERS = [
    "A", "B", "C", "D", "E", "F", "G", "H",
    "I", "J", "K", "L", "M", "N", "O", "P",
]
# F:I (Activity ID, Activity Name, Start, Finish) are read from the source
# but are never included in the normalized document table.
INCLUDED_COLUMN_INDEXES = [0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15]
DOCUMENT_PROGRESS = 0.05
PUBLICATION_STATUS = "Not Published"

COLUMNS = [
    "Serial No.",
    "Discipline",
    "Area",
    "Document No.",
    "Document Title",
    "Listed in Contract",
    "Document Action",
    "Required MH",
    "Plan Issue Date",
    "Deliverable Type",
    "Responsible",
    "Remarks",
]

TEXT_COLUMNS = [
    "Discipline",
    "Area",
    "Document No.",
    "Document Title",
    "Listed in Contract",
    "Document Action",
    "Deliverable Type",
    "Responsible",
    "Remarks",
]

DOCUMENT_LIST_COLUMNS = [
    "Discipline",
    "Area",
    "Document No.",
    "Document Title",
    "Plan Issue Date",
    "Document Progress",
    "Publication Status",
    "Document Action",
    "Deliverable Type",
    "Responsible",
    "Listed in Contract",
    "Assignment",
    "Remarks",
]


class WorkbookSchemaError(ValueError):
    """Raised when the Excel workbook does not match the expected TKPR2 template."""


def _header_text(raw: pd.DataFrame, column_index: int) -> str:
    values = raw.iloc[:HEADER_ROW_COUNT, column_index].tolist()
    return " ".join(str(value).strip() for value in values if pd.notna(value)).lower()


def _clean_text(value: object) -> object:
    if pd.isna(value):
        return pd.NA
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "<na>"}:
        return pd.NA
    return text


def parse_excel_date(value: object) -> pd.Timestamp | pd.NaT:
    """Safely convert an Excel date value to a pandas Timestamp.

    Excel serial dates use the 1899-12-30 epoch. Native date cells and
    day/month/year text values are also supported.
    """

    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value.normalize()
    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value).normalize()
    if isinstance(value, (int, float)):
        try:
            return (pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")).normalize()
        except (OverflowError, ValueError, TypeError):
            return pd.NaT

    text = str(value).strip()
    if not text:
        return pd.NaT
    try:
        numeric = float(text.replace(",", "."))
        if numeric > 10_000:
            return (pd.Timestamp("1899-12-30") + pd.to_timedelta(numeric, unit="D")).normalize()
    except ValueError:
        pass
    return pd.to_datetime(text, dayfirst=True, errors="coerce")


def _validate_headers(raw: pd.DataFrame) -> None:
    if raw.shape[1] != len(SOURCE_COLUMN_LETTERS):
        raise WorkbookSchemaError(
            f"Range A:P is expected in sheet '{SHEET_NAME}'; {raw.shape[1]} columns were read."
        )

    expected_tokens = {
        1: "discipline",         # B / index 1
        3: "doc. no",            # D / index 3
        4: "doc. title",         # E / index 4
        5: "activity id",        # F / index 5; under the merged PMS header
        6: "activity name",      # G / index 6
        7: "start",              # H / index 7
        8: "finish",             # I / index 8
        12: "plan issue date",   # M / index 12
        13: "deliverable type",  # N / index 13
        14: "issued to",         # O / index 14
    }
    mismatches = []
    for column_index, token in expected_tokens.items():
        actual = _header_text(raw, column_index)
        if token not in actual:
            excel_letter = SOURCE_COLUMN_LETTERS[column_index]
            mismatches.append(f"{excel_letter}: expected '{token}', found '{actual}'")

    if mismatches:
        raise WorkbookSchemaError(
            "Excel headers do not match the expected TKPR2 template: " + "; ".join(mismatches)
        )


def read_document_master(source: str | Path | BinaryIO) -> pd.DataFrame:
    """Convert the current Document Master List into a normalized table."""

    if hasattr(source, "seek"):
        source.seek(0)

    try:
        raw = pd.read_excel(
            source,
            sheet_name=SHEET_NAME,
            header=None,
            usecols=READ_USECOLS,
            nrows=SOURCE_MAX_ROWS,
            engine="openpyxl",
        )
    except ValueError as exc:
        raise WorkbookSchemaError(
            f"Sheet '{SHEET_NAME}' could not be read. Check the file and sheet name."
        ) from exc

    _validate_headers(raw)

    data = raw.iloc[HEADER_ROW_COUNT:, INCLUDED_COLUMN_INDEXES].copy()
    data.columns = COLUMNS
    data.insert(0, "Excel Row", data.index + 1)

    # Keep actual document records; discard blank and formatting-only rows.
    data = data.loc[data["Document No."].notna() | data["Document Title"].notna()].copy()

    for column in TEXT_COLUMNS:
        data[column] = data[column].map(_clean_text).astype("string")

    data["Serial No."] = pd.to_numeric(data["Serial No."], errors="coerce").astype("Int64")
    data["Required MH"] = pd.to_numeric(data["Required MH"], errors="coerce")

    # Critical mapping: column M is assigned directly to "Plan Issue Date".
    data["Plan Issue Date"] = data["Plan Issue Date"].map(parse_excel_date)
    data["Listed in Contract"] = data["Listed in Contract"].fillna("Blank / Not specified")
    for column in ["Discipline", "Area", "Document Action", "Deliverable Type", "Responsible"]:
        data[column] = data[column].fillna("(Blank)")

    data["Responsible"] = data["Responsible"].str.upper()
    data["Deliverable Type"] = data["Deliverable Type"].str.title()
    data["Document Action"] = data["Document Action"].str.title()

    data["Is Placeholder"] = data["Document No."].fillna("").str.contains(
        r"X{3,}", case=False, regex=True
    )
    data["Plan Month"] = data["Plan Issue Date"].dt.to_period("M").dt.to_timestamp()
    data["Assignment"] = data.apply(_derive_assignment, axis=1)
    # User-confirmed rule: all documents are not yet published and each document
    # is displayed with a fixed progress value of 5%.
    data["Document Progress"] = DOCUMENT_PROGRESS
    data["Publication Status"] = PUBLICATION_STATUS

    return data.reset_index(drop=True)


def _derive_assignment(row: pd.Series) -> str:
    deliverable_type = str(row["Deliverable Type"]).strip()
    responsible = str(row["Responsible"]).strip()
    if deliverable_type in {"", "(Blank)", "<NA>"} or responsible in {"", "(Blank)", "<NA>"}:
        return "Unassigned"
    return f"{deliverable_type} / {responsible}"


def build_quality_issues(data: pd.DataFrame) -> pd.DataFrame:
    """List schema and data-consistency issues in the source rows."""

    issues: list[dict[str, object]] = []

    def add(row: pd.Series, severity: str, issue: str) -> None:
        issues.append(
            {
                "Severity": severity,
                "Excel Row": int(row["Excel Row"]),
                "Document No.": row["Document No."],
                "Document Title": row["Document Title"],
                "Issue": issue,
            }
        )

    required = [
        "Discipline",
        "Area",
        "Document No.",
        "Document Title",
        "Document Action",
        "Plan Issue Date",
        "Deliverable Type",
        "Responsible",
    ]

    for _, row in data.iterrows():
        for column in required:
            value = row[column]
            if pd.isna(value) or value in {"", "(Blank)"}:
                add(row, "Error", f"Required field is blank: {column}")

        if pd.isna(row["Serial No."]):
            add(row, "Warning", "Serial No. is blank")

        if row["Listed in Contract"] == "Blank / Not specified":
            add(row, "Warning", "Listed in Contract is blank")

        if bool(row["Is Placeholder"]):
            add(row, "Info", "Document number contains an XXXX placeholder")

    real_docs = data.loc[~data["Is Placeholder"] & data["Document No."].notna()]
    duplicated = real_docs[real_docs["Document No."].duplicated(keep=False)]
    for _, row in duplicated.iterrows():
        add(row, "Warning", "Duplicate non-placeholder document number")

    columns = ["Severity", "Excel Row", "Document No.", "Document Title", "Issue"]
    return pd.DataFrame(issues, columns=columns)


def export_columns() -> list[str]:
    """Return the primary Excel fields used for display and export."""

    result = ["Excel Row", *COLUMNS]
    plan_index = result.index("Plan Issue Date") + 1
    result[plan_index:plan_index] = ["Document Progress", "Publication Status"]
    return result


st.set_page_config(
    page_title="TKPR2 Document Master List",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      :root {
        color-scheme:light;
        --ink:#13283F; --muted:#526A82;
        --navy:#0B3A63; --deep-navy:#082E4F; --blue:#1E6FAE;
        --teal:#0B8F92; --orange:#E67E22;
        --line:#D2DEE9; --grid:#DFE7EF; --page:#F3F6FA;
        --card:#FFFFFF; --soft-blue:#E9F2FA; --soft-teal:#E8F6F5;
        --shadow:0 10px 28px rgba(24,53,82,.09);
      }
      html, body, [class*="st-"] {
        font-family:"Segoe UI", Inter, Arial, sans-serif;
      }
      .stApp {
        color:var(--ink);
        background:
          radial-gradient(circle at 92% 0%, rgba(30,111,174,.10), transparent 28rem),
          linear-gradient(180deg, #F8FAFC 0%, var(--page) 55%, #EDF2F7 100%);
      }
      [data-testid="stAppViewContainer"] { background:transparent; }
      [data-testid="stHeader"] {
        background:rgba(248,250,252,.94);
        border-bottom:1px solid rgba(210,222,233,.85);
        backdrop-filter:blur(12px);
      }
      [data-testid="stSidebar"] {
        background:linear-gradient(180deg, #EDF3F8 0%, #E5EDF5 100%);
        border-right:1px solid #CBD8E4;
      }
      [data-testid="stSidebar"] h1,
      [data-testid="stSidebar"] h2,
      [data-testid="stSidebar"] h3,
      [data-testid="stSidebar"] p,
      [data-testid="stSidebar"] label { color:var(--ink); }
      [data-testid="stSidebar"] [data-baseweb="select"] > div,
      [data-testid="stSidebar"] [data-baseweb="input"] > div,
      [data-testid="stSidebar"] input {
        background:#FFFFFF !important;
        color:var(--ink) !important;
        border-color:#C4D3E0 !important;
      }
      [data-testid="stSidebar"] [data-baseweb="tag"] {
        background:#D9EAF5 !important;
        border:1px solid #BCD4E5 !important;
      }
      [data-testid="stSidebar"] [data-baseweb="tag"] span {
        color:#173A57 !important;
      }
      [data-testid="stAppDeployButton"], #MainMenu {
        display:none !important;
      }
      /* Keep sidebar show/hide controls reachable at all times. */
      [data-testid="stSidebarCollapseButton"],
      [data-testid="stSidebarCollapsedControl"] {
        display:flex !important;
        visibility:visible !important;
        opacity:1 !important;
        pointer-events:auto !important;
        z-index:1000000 !important;
      }
      [data-testid="stSidebarCollapseButton"] button,
      [data-testid="stSidebarCollapsedControl"] button {
        display:flex !important;
        visibility:visible !important;
        opacity:1 !important;
        background:#FFFFFF !important;
        color:transparent !important;
        font-size:0 !important;
        border:1px solid #AFC7D8 !important;
        border-radius:9px !important;
        box-shadow:0 3px 9px rgba(24,53,82,.14) !important;
        min-width:34px !important;
        width:34px !important;
        height:32px !important;
        align-items:center !important;
        justify-content:center !important;
      }
      [data-testid="stSidebarCollapseButton"] button > *,
      [data-testid="stSidebarCollapsedControl"] button > * {
        display:none !important;
      }
      [data-testid="stSidebarCollapseButton"] button::after {
        content:"<<" !important;
        display:block !important;
        color:#0B3A63 !important;
        font-size:19px !important;
        line-height:1 !important;
        font-weight:900 !important;
        letter-spacing:-3px !important;
      }
      [data-testid="stSidebarCollapsedControl"] button::after {
        content:">>" !important;
        display:block !important;
        color:#0B3A63 !important;
        font-size:19px !important;
        line-height:1 !important;
        font-weight:900 !important;
        letter-spacing:-3px !important;
      }
      [data-testid="stSidebarCollapsedControl"] {
        position:fixed !important;
        top:.75rem !important;
        left:.65rem !important;
      }
      .block-container {
        max-width:1680px;
        padding-top:2.75rem !important;
        padding-bottom:.75rem !important;
      }
      .block-container > [data-testid="stVerticalBlock"],
      [data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] {
        gap:1rem !important;
      }
      [data-testid="stHorizontalBlock"] {
        gap:1rem !important;
      }
      /* Keep title and REFRESH below Streamlit fixed header. */
      [data-testid="stMainBlockContainer"] {
        padding-top:2.75rem !important;
      }
      [data-testid="stMainBlockContainer"] [data-testid="stHorizontalBlock"]:first-of-type {
        position:relative !important;
        z-index:5 !important;
        min-height:58px !important;
        overflow:visible !important;
      }
      [data-testid="stMainBlockContainer"] [data-testid="stHorizontalBlock"]:first-of-type h1 {
        margin:0 !important;
        padding:0 !important;
        line-height:1.15 !important;
        overflow:visible !important;
      }
      [data-testid="stMainBlockContainer"] [data-testid="stHorizontalBlock"]:first-of-type [data-testid="stButton"] {
        margin:0 !important;
        padding:0 !important;
      }
      h1, h2, h3 {
        color:var(--navy) !important;
        letter-spacing:-.02em;
      }
      h1 {
        font-size:2.25rem !important;
        font-weight:800 !important;
        margin-top:0 !important;
        margin-bottom:.25rem !important;
        padding-top:0 !important;
      }
      h2 { font-weight:750 !important; }
      h3 {
        font-weight:700 !important;
        margin-top:.35rem !important;
        margin-bottom:.5rem !important;
      }
      p, li, label, [data-testid="stCaptionContainer"] {
        color:var(--ink);
      }
      [data-testid="stCaptionContainer"] {
        background:#E9F1F8;
        border:1px solid #D0DFEB;
        border-radius:10px;
        padding:.65rem .8rem;
        font-size:.88rem;
      }
      .stHeadingWithActionElements a { display:none !important; }
      [data-testid="stMetric"] {
        background:linear-gradient(145deg, #FFFFFF 0%, #F8FBFD 100%);
        border:1px solid var(--line);
        border-top:4px solid var(--blue);
        border-radius:14px;
        padding:13px 16px 15px !important;
        box-shadow:var(--shadow);
        height: 128px; min-height: 128px; box-sizing: border-box;
        display:flex; flex-direction:column; justify-content:flex-start;
        align-items:stretch; overflow:hidden;
        transition:transform .18s ease, box-shadow .18s ease;
      }
      [data-testid="stMetric"]:hover {
        transform:translateY(-2px);
        box-shadow:0 14px 30px rgba(24,53,82,.13);
      }
      [data-testid="stMetricLabel"] {
        color:#405A73 !important; font-size:14px !important; line-height:20px !important;
        font-weight:700; min-height:20px; height:20px;
        display:flex; align-items:center; margin:0 !important;
      }
      [data-testid="stMetricLabel"] p {
        color:#405A73 !important;
        font-size:14px !important; line-height:20px !important;
        margin:0 !important; white-space:nowrap;
      }
      [data-testid="stMetricValue"] {
        color:var(--deep-navy) !important; font-size:32px !important; line-height:40px !important;
        font-weight:750; min-height:40px; height:40px; margin-top:8px !important;
        display:flex; align-items:center;
      }
      [data-testid="stMetricValue"] > div {
        color:var(--deep-navy) !important;
        font-size:32px !important; line-height:40px !important;
      }
      [data-testid="stMetricDelta"] {
        color:#36536D !important;
        font-size:12px !important; line-height:16px !important;
        min-height:16px; margin-top:4px !important;
      }
      [data-testid="stProgress"] {
        background:#FFFFFF;
        border:1px solid var(--line);
        border-radius:12px;
        padding:.7rem .85rem .65rem;
        margin:0 !important;
        box-shadow:0 4px 14px rgba(24,53,82,.06);
      }
      [data-testid="stProgress"] p {
        color:#304D68 !important;
        font-weight:650;
      }
      [data-testid="stProgressBar"] > div { background:#D8E3EC; }
      [data-testid="stProgressBar"] > div > div {
        background:linear-gradient(90deg, var(--blue), var(--teal)) !important;
      }
      /* Three clearly separated button-style tabs. */
      [data-testid="stTabs"] [role="tablist"],
      [data-baseweb="tab-list"] {
        display:flex !important;
        width:min(720px, 100%) !important;
        gap:1rem !important;
        padding:.35rem 0 !important;
        margin:.25rem 0 .9rem 1.25rem !important;
        background:transparent !important;
        border:0 !important;
        box-shadow:none !important;
        overflow:visible !important;
      }
      [data-testid="stTabs"] button[role="tab"],
      [data-baseweb="tab"] {
        flex:1 1 0 !important;
        width:auto !important;
        min-width:150px !important;
        height:2.85rem !important;
        min-height:2.85rem !important;
        justify-content:center !important;
        padding:0 1.15rem !important;
        margin:0 !important;
        border:1px solid #AFC7D8 !important;
        border-radius:10px !important;
        background:#FFFFFF !important;
        color:#0B3A63 !important;
        font-weight:750 !important;
        opacity:1 !important;
        box-shadow:0 3px 9px rgba(24,53,82,.10) !important;
        transition:all .16s ease !important;
      }
      [data-testid="stTabs"] button[role="tab"] p,
      [data-testid="stTabs"] button[role="tab"] span,
      [data-baseweb="tab"] p,
      [data-baseweb="tab"] span {
        color:#0B3A63 !important;
        font-size:1.02rem !important;
        font-weight:800 !important;
        letter-spacing:.01em !important;
        white-space:nowrap !important;
      }
      [data-testid="stTabs"] button[role="tab"]:hover,
      [data-baseweb="tab"]:hover {
        background:#E9F5FB !important;
        border-color:#4F91B5 !important;
        transform:translateY(-1px) !important;
        box-shadow:0 5px 13px rgba(24,92,130,.16) !important;
      }
      [data-testid="stTabs"] button[role="tab"][aria-selected="true"],
      [data-baseweb="tab"][aria-selected="true"] {
        background:linear-gradient(135deg, #1267A5 0%, #0B8F92 100%) !important;
        border-color:#0B6F86 !important;
        box-shadow:0 5px 14px rgba(11,111,134,.22) !important;
      }
      [data-testid="stTabs"] button[role="tab"][aria-selected="true"] p,
      [data-testid="stTabs"] button[role="tab"][aria-selected="true"] span,
      [data-baseweb="tab"][aria-selected="true"] p,
      [data-baseweb="tab"][aria-selected="true"] span {
        color:#FFFFFF !important;
        font-weight:800 !important;
      }
      [data-testid="stTabs"] [data-baseweb="tab-highlight"],
      [data-testid="stTabs"] [data-baseweb="tab-border"],
      [data-baseweb="tab-highlight"],
      [data-baseweb="tab-border"] {
        display:none !important;
      }
      /* Clickable KPI cards that navigate to dedicated list views. */
      .kpi-nav-card {
        display:flex; flex-direction:column; min-height:122px; padding:14px 16px 12px;
        background:#FFFFFF; border:1px solid #C7D9E6; border-top:4px solid #1776B6;
        border-radius:14px; box-shadow:0 5px 14px rgba(24,53,82,.08);
        text-decoration:none !important; color:#082E4F !important;
        transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease;
      }
      .kpi-nav-card:hover {
        transform:translateY(-2px); border-color:#6EA9C9;
        box-shadow:0 8px 19px rgba(24,92,130,.15);
      }
      .kpi-nav-label { font-size:13px; font-weight:700; color:#284E6B; }
      .kpi-nav-value { font-size:31px; line-height:1.15; font-weight:850; margin-top:9px; color:#082E4F; }
      .kpi-nav-note { font-size:11px; font-weight:650; color:#507089; margin-top:7px; }
      [data-testid="stPlotlyChart"] {
        background:var(--card); border:1px solid var(--line); border-radius:15px;
        padding:10px; box-shadow:var(--shadow);
        margin:0 !important;
        overflow:hidden;
      }
      [data-testid="stDataFrame"] {
        background:#FFFFFF;
        border:1px solid var(--line);
        border-radius:14px;
        padding:.35rem;
        margin:0 !important;
        box-shadow:var(--shadow);
        overflow:hidden;
      }
      .quality-ok { padding:12px; border-radius:9px; background:#E8F6F0; color:#17653A; }
      .small-note { color:var(--muted); font-size:.86rem; }
      .stButton > button,
      div[data-testid="stDownloadButton"] button {
        border-radius:9px !important;
        border:1px solid #B8CBDB !important;
        background:#FFFFFF !important;
        color:var(--navy) !important;
        font-weight:700 !important;
        box-shadow:0 3px 10px rgba(24,53,82,.08);
      }
      .stButton > button:hover,
      div[data-testid="stDownloadButton"] button:hover {
        border-color:var(--blue) !important;
        background:#EAF3FB !important;
        color:var(--deep-navy) !important;
      }
      .stButton > button[kind="primary"] {
        background:linear-gradient(135deg, var(--navy), var(--blue)) !important;
        color:#FFFFFF !important;
        border-color:var(--navy) !important;
      }
      /* Full-size, readable REFRESH control in the title row. */
      div[data-testid="stButton"] button[kind="primary"],
      div[data-testid="stButton"] button[data-testid="stBaseButton-primary"] {
        width:100% !important;
        min-width:132px !important;
        height:48px !important;
        min-height:48px !important;
        padding:0 1.15rem !important;
        border:1px solid #0A4F78 !important;
        border-radius:11px !important;
        background:linear-gradient(135deg, #1E6FAE 0%, #0B8F92 100%) !important;
        color:#FFFFFF !important;
        box-shadow:0 6px 16px rgba(11,92,125,.22) !important;
        display:flex !important;
        align-items:center !important;
        justify-content:center !important;
        gap:.4rem !important;
        white-space:nowrap !important;
        overflow:visible !important;
        line-height:1 !important;
      }
      div[data-testid="stButton"] button[kind="primary"] p,
      div[data-testid="stButton"] button[kind="primary"] span,
      div[data-testid="stButton"] button[data-testid="stBaseButton-primary"] p,
      div[data-testid="stButton"] button[data-testid="stBaseButton-primary"] span {
        color:#FFFFFF !important;
        font-size:.93rem !important;
        font-weight:800 !important;
        line-height:1 !important;
        margin:0 !important;
        white-space:nowrap !important;
        text-align:center !important;
      }
      div[data-testid="stButton"] button[kind="primary"]:hover,
      div[data-testid="stButton"] button[data-testid="stBaseButton-primary"]:hover {
        background:linear-gradient(135deg, #155E91 0%, #087F82 100%) !important;
        border-color:#063B5E !important;
        box-shadow:0 8px 19px rgba(11,92,125,.29) !important;
        transform:translateY(-1px) !important;
      }
      div[data-testid="stButton"] button[kind="primary"]:focus-visible,
      div[data-testid="stButton"] button[data-testid="stBaseButton-primary"]:focus-visible {
        outline:3px solid rgba(11,143,146,.30) !important;
        outline-offset:2px !important;
      }
      div[data-testid="stDownloadButton"] button { width:100%; }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False, max_entries=12)
def build_export_xlsx(frame: pd.DataFrame) -> bytes:
    """Create the current Document List as a real XLSX with openpyxl only.

    This deliberately avoids pandas' xlsxwriter engine, so downloading an Excel
    file cannot fail merely because the optional xlsxwriter package is absent.
    """

    output = BytesIO()
    export = frame[DOCUMENT_LIST_COLUMNS].copy()
    export = export.rename(columns={"Responsible": "Parner"})

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Document List"

    worksheet.append(export.columns.tolist())
    for row in export.itertuples(index=False, name=None):
        excel_row = []
        for value in row:
            if value is None or pd.isna(value):
                excel_row.append(None)
            elif isinstance(value, pd.Timestamp):
                excel_row.append(value.to_pydatetime())
            elif hasattr(value, "item"):
                excel_row.append(value.item())
            else:
                excel_row.append(value)
        worksheet.append(excel_row)

    header_fill = PatternFill("solid", fgColor="0B2E59")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.row_dimensions[1].height = 34

    for col_index, column in enumerate(export.columns, start=1):
        sample = export[column].dropna().astype(str).head(300)
        sample_width = int(sample.str.len().max()) if not sample.empty else 0
        worksheet.column_dimensions[get_column_letter(col_index)].width = min(
            max(len(column) + 2, sample_width + 2, 11), 42
        )
        if column == "Plan Issue Date":
            for cell in worksheet.iter_cols(
                min_col=col_index, max_col=col_index, min_row=2, max_row=worksheet.max_row
            ):
                for item in cell:
                    item.number_format = "dd.mm.yyyy"
        elif column == "Document Progress":
            worksheet.column_dimensions[get_column_letter(col_index)].width = 18
            for cell in worksheet.iter_cols(
                min_col=col_index, max_col=col_index, min_row=2, max_row=worksheet.max_row
            ):
                for item in cell:
                    item.number_format = "0%"

    workbook.save(output)
    return output.getvalue()


def render_document_list_table(frame: pd.DataFrame) -> None:
    """Render an in-table searchable and sortable Document List."""

    table_frame = frame[DOCUMENT_LIST_COLUMNS].copy()
    table_frame["Plan Issue Date"] = pd.to_datetime(
        table_frame["Plan Issue Date"], errors="coerce"
    )

    labels = {column: column for column in DOCUMENT_LIST_COLUMNS}
    labels["Responsible"] = "Parner"
    widths = {
        "Discipline": 120,
        "Area": 160,
        "Document No.": 280,
        "Document Title": 420,
        "Plan Issue Date": 150,
        "Document Progress": 150,
        "Publication Status": 160,
        "Document Action": 150,
        "Deliverable Type": 150,
        "Responsible": 120,
        "Listed in Contract": 165,
        "Assignment": 170,
        "Remarks": 300,
    }

    rows: list[dict[str, dict[str, object]]] = []
    for _, row in table_frame.iterrows():
        output_row: dict[str, dict[str, object]] = {}
        for column in DOCUMENT_LIST_COLUMNS:
            value = row[column]
            if pd.isna(value):
                display_value = ""
                sort_value: object = None
            elif column == "Plan Issue Date":
                timestamp = pd.Timestamp(value)
                display_value = timestamp.strftime("%d.%m.%Y")
                sort_value = int(timestamp.timestamp())
            elif column == "Document Progress":
                display_value = f"{float(value):.0%}"
                sort_value = float(value)
            else:
                display_value = str(value)
                sort_value = display_value.casefold()
            output_row[column] = {"display": display_value, "sort": sort_value}
        rows.append(output_row)

    safe_rows_json = json.dumps(rows, ensure_ascii=False).replace("</", "<\\/")
    column_specs = [
        {"key": column, "label": labels[column], "width": widths[column]}
        for column in DOCUMENT_LIST_COLUMNS
    ]
    safe_columns_json = json.dumps(column_specs, ensure_ascii=False).replace("</", "<\\/")

    table_html = f"""
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <style>
        :root {{
          color-scheme:light;
          --ink:#13283F; --muted:#526A82; --navy:#0B3A63;
          --deep:#082E4F; --blue:#1E6FAE; --teal:#0B8F92;
          --line:#CFDCE7; --grid:#DEE7EF; --page:#FFFFFF; --head:#0B3A63;
        }}
        * {{ box-sizing:border-box; }}
        html, body {{
          margin:0; padding:0; background:transparent;
          font-family:"Segoe UI", Inter, Arial, sans-serif; color:var(--ink);
        }}
        .table-card {{
          border:1px solid #BDD0DF; border-radius:14px; background:var(--page);
          overflow:hidden; box-shadow:0 10px 26px rgba(24,53,82,.10);
        }}
        .toolbar {{
          display:flex; justify-content:space-between; align-items:center; gap:16px; padding:11px 14px;
          background:linear-gradient(90deg, #EAF2F8 0%, #F5F8FB 100%);
          border-bottom:1px solid var(--line); font-size:12px; color:#425D76;
        }}
        .hint {{ font-weight:650; color:#284A68; }}
        #rowCount {{
          flex:0 0 auto; border-radius:999px; padding:4px 9px;
          background:#DDECF6; color:var(--navy); font-weight:750;
        }}
        .table-wrap {{ height:570px; overflow:auto; }}
        table {{ border-collapse:separate; border-spacing:0; min-width:2495px; width:100%; table-layout:fixed; }}
        thead {{ position:sticky; top:0; z-index:5; }}
        th {{
          background:linear-gradient(180deg, #114D7E 0%, var(--head) 100%);
          border-right:1px solid rgba(255,255,255,.17); border-bottom:1px solid #062D4D;
          padding:9px 9px 10px; text-align:left; vertical-align:top;
        }}
        th:last-child, td:last-child {{ border-right:none; }}
        .sort-button {{
          width:100%; border:0; background:transparent; padding:0 0 7px; text-align:left;
          color:#FFFFFF; font-size:12px; font-weight:750; cursor:pointer; white-space:nowrap;
        }}
        .sort-button:hover {{ color:#D7F3F3; }}
        .sort-mark {{ float:right; color:#75D7D4; }}
        .column-search {{
          width:100%; height:33px; border:1px solid #AEC5D8; border-radius:7px;
          background:#FFFFFF; color:var(--ink); padding:6px 8px; font-size:12px; outline:none;
        }}
        .column-search::placeholder {{ color:#75889A; opacity:1; }}
        .column-search:focus {{ border-color:#43C3C2; box-shadow:0 0 0 3px rgba(67,195,194,.22); }}
        tbody tr {{ background:#FFFFFF; }}
        tbody tr:nth-child(even) {{ background:#EDF4F9; }}
        tbody tr:hover {{ background:#DDF1F0; }}
        td {{
          border-right:1px solid var(--grid); border-bottom:1px solid var(--grid);
          padding:9px 8px; font-size:12px; line-height:18px; vertical-align:top;
          color:var(--ink); white-space:normal; overflow-wrap:anywhere;
        }}
        .empty {{ padding:30px; text-align:center; color:#526A82; background:#F8FAFC; }}
      </style>
    </head>
    <body>
      <div class="table-card">
        <div class="toolbar">
          <span class="hint">Use commas, semicolons, or new lines for AND terms in a column. All terms must exist in the same cell. Click a title to sort.</span>
          <span id="rowCount"></span>
        </div>
        <div class="table-wrap">
          <table id="documentTable">
            <colgroup id="colGroup"></colgroup>
            <thead><tr id="headerRow"></tr></thead>
            <tbody id="tableBody"></tbody>
          </table>
        </div>
      </div>
      <script>
        const allRows = {safe_rows_json};
        const columns = {safe_columns_json};
        const filters = Object.fromEntries(columns.map(c => [c.key, ""]));
        let sortKey = null;
        let sortDirection = 1;

        const colGroup = document.getElementById("colGroup");
        const headerRow = document.getElementById("headerRow");
        const tableBody = document.getElementById("tableBody");
        const rowCount = document.getElementById("rowCount");

        function escapeHtml(value) {{
          return String(value ?? "")
            .replaceAll("&", "&amp;").replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;").replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
        }}

        function buildHeader() {{
          columns.forEach(column => {{
            const col = document.createElement("col");
            col.style.width = `${{column.width}}px`;
            colGroup.appendChild(col);

            const th = document.createElement("th");
            const button = document.createElement("button");
            button.className = "sort-button";
            button.dataset.key = column.key;
            button.innerHTML = `<span>${{escapeHtml(column.label)}}</span><span class="sort-mark"></span>`;
            button.addEventListener("click", () => setSort(column.key));

            const input = document.createElement("input");
            input.className = "column-search";
            input.type = "search";
            input.placeholder = "Search term1, term2...";
            input.setAttribute("aria-label", `Search ${{column.label}}`);
            input.addEventListener("click", event => event.stopPropagation());
            input.addEventListener("input", event => {{
              filters[column.key] = event.target.value;
              renderRows();
            }});
            th.append(button, input);
            headerRow.appendChild(th);
          }});
        }}

        function setSort(key) {{
          if (sortKey === key) sortDirection *= -1;
          else {{ sortKey = key; sortDirection = 1; }}
          document.querySelectorAll(".sort-mark").forEach(mark => mark.textContent = "");
          const active = document.querySelector(`.sort-button[data-key="${{CSS.escape(key)}}"] .sort-mark`);
          if (active) active.textContent = sortDirection === 1 ? "▲" : "▼";
          renderRows();
        }}

        function normalizeSearch(value) {{
          return String(value ?? "")
            .normalize("NFKD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLocaleLowerCase();
        }}

        function searchTerms(value) {{
          return String(value ?? "")
            .split(/[,;\\n]+/u)
            .map(term => normalizeSearch(term.trim()))
            .filter(Boolean);
        }}

        function filteredRows() {{
          const result = allRows.filter(row => columns.every(column => {{
            const terms = searchTerms(filters[column.key]);
            if (!terms.length) return true;
            const cellText = normalizeSearch(row[column.key].display);
            return terms.every(term => cellText.includes(term));
          }}));
          if (!sortKey) return result;
          return result.sort((a, b) => {{
            const av = a[sortKey].sort;
            const bv = b[sortKey].sort;
            if (av === null || av === undefined || av === "") return 1;
            if (bv === null || bv === undefined || bv === "") return -1;
            if (av < bv) return -1 * sortDirection;
            if (av > bv) return 1 * sortDirection;
            return 0;
          }});
        }}

        function renderRows() {{
          const rows = filteredRows();
          rowCount.textContent = `${{rows.length.toLocaleString()}} of ${{allRows.length.toLocaleString()}} records`;
          if (!rows.length) {{
            tableBody.innerHTML = `<tr><td class="empty" colspan="${{columns.length}}">No records match the table filters.</td></tr>`;
            return;
          }}
          tableBody.innerHTML = rows.map(row =>
            `<tr>${{columns.map(column => `<td>${{escapeHtml(row[column.key].display)}}</td>`).join("")}}</tr>`
          ).join("");
        }}

        buildHeader();
        renderRows();
      </script>
    </body>
    </html>
    """
    st.iframe(
        table_html,
        width="stretch",
        height=655,
        alt="Searchable and sortable Document List table",
    )


def render_area_summary_table(frame: pd.DataFrame) -> None:
    """Render the Area summary with a light, high-contrast dashboard style."""

    headers = [
        ("Area", "Area"),
        ("Documents", "Documents"),
        ("RET", "RET"),
        ("TOYO", "TOYO"),
        ("External", "External"),
        ("Internal", "Internal"),
        ("Document_Progress", "Document Progress"),
    ]
    header_html = "".join(
        f'<th class="{"area-col" if key == "Area" else "number-col"}">{html.escape(label)}</th>'
        for key, label in headers
    )

    body_rows: list[str] = []
    for _, row in frame.iterrows():
        area = "" if pd.isna(row["Area"]) else html.escape(str(row["Area"]))
        progress = 0.0 if pd.isna(row["Document_Progress"]) else float(row["Document_Progress"])
        progress = max(0.0, min(progress, 1.0))
        progress_percent = progress * 100
        numeric_cells = "".join(
            f'<td class="number-col">{int(row[column]):,}</td>'
            for column in ["Documents", "RET", "TOYO", "External", "Internal"]
        )
        body_rows.append(
            "<tr>"
            f'<td class="area-col">{area}</td>'
            f"{numeric_cells}"
            '<td class="progress-col">'
            '<div class="progress-wrap">'
            '<div class="progress-track" aria-hidden="true">'
            f'<div class="progress-fill" style="width:{progress_percent:.1f}%"></div>'
            '</div>'
            f'<span class="progress-value">{progress_percent:.0f}%</span>'
            '</div>'
            '</td>'
            "</tr>"
        )

    table_html = f"""
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <style>
        :root {{ color-scheme:light; }}
        * {{ box-sizing:border-box; }}
        html, body {{
          margin:0; padding:0; background:transparent; color:#13283F;
          font-family:"Segoe UI", Inter, Arial, sans-serif;
        }}
        .summary-card {{
          width:82%; max-width:1120px; margin:0 auto; overflow:hidden;
          background:#FFFFFF; border:1px solid #C8D8E5;
          border-radius:14px; box-shadow:0 8px 22px rgba(24,53,82,.09);
        }}
        table {{ width:100%; border-collapse:separate; border-spacing:0; table-layout:fixed; }}
        thead th {{
          height:40px; padding:8px 9px; background:linear-gradient(180deg, #155E91 0%, #0B3A63 100%);
          color:#FFFFFF; border-right:1px solid rgba(255,255,255,.18);
          border-bottom:1px solid #082E4F; font-size:13px; line-height:18px;
          font-weight:750; text-align:center; vertical-align:middle;
        }}
        thead th:last-child, tbody td:last-child {{ border-right:none; }}
        tbody tr {{ background:#F8FBFD; }}
        tbody tr:nth-child(even) {{ background:#EAF4F8; }}
        tbody tr:hover {{ background:#DDF1F0; }}
        tbody td {{
          height:38px; padding:7px 9px; border-right:1px solid #D7E4ED;
          border-bottom:1px solid #D7E4ED; color:#173A57; font-size:13px;
          line-height:18px; font-weight:600; text-align:center; vertical-align:middle;
        }}
        tbody tr:last-child td {{ border-bottom:none; }}
        .area-col {{ width:20%; text-align:left !important; }}
        .number-col {{ width:9.5%; text-align:center !important; font-variant-numeric:tabular-nums; }}
        .progress-col {{ width:22.5%; text-align:center !important; }}
        .progress-wrap {{ display:flex; align-items:center; justify-content:center; gap:10px; width:100%; }}
        .progress-track {{
          flex:1 1 auto; height:9px; min-width:75px; overflow:hidden;
          background:#D9E8F2; border:1px solid #C4D8E6; border-radius:999px;
          box-shadow:inset 0 1px 2px rgba(24,53,82,.06);
        }}
        .progress-fill {{
          height:100%; min-width:6px; border-radius:999px;
          background:linear-gradient(90deg, #1E6FAE 0%, #0B8F92 100%);
        }}
        .progress-value {{
          flex:0 0 38px; color:#0B3A63; font-weight:750; text-align:right;
          font-variant-numeric:tabular-nums;
        }}
      </style>
    </head>
    <body>
      <div class="summary-card">
        <table aria-label="Summary by Area">
          <thead><tr>{header_html}</tr></thead>
          <tbody>{''.join(body_rows)}</tbody>
        </table>
      </div>
    </body>
    </html>
    """
    table_height = min(470, 44 + max(len(frame), 1) * 38 + 4)
    st.iframe(
        table_html,
        width="stretch",
        height=table_height,
        alt="Colorful Summary by Area table",
    )


def fmt_int(value: int) -> str:
    return f"{int(value):,}".replace(",", ".")


def normalised_options(series: pd.Series) -> list[str]:
    return sorted(str(value) for value in series.dropna().unique())


def filter_frame(
    frame: pd.DataFrame,
    disciplines: list[str],
    areas: list[str],
    responsible: list[str],
    deliverable_types: list[str],
    actions: list[str],
    search_text: str,
) -> pd.DataFrame:
    filtered = frame.copy()
    optional_filters = {
        "Discipline": disciplines,
        "Area": areas,
        "Responsible": responsible,
        "Deliverable Type": deliverable_types,
        "Document Action": actions,
    }
    for column, selected_values in optional_filters.items():
        # An empty selection means "All"; filter only when values are selected.
        if selected_values:
            filtered = filtered[filtered[column].isin(selected_values)]

    terms = [term.strip() for term in search_text.replace(";", " ").split() if term.strip()]
    for term in terms:
        mask = (
            filtered["Document No."].fillna("").str.contains(term, case=False, regex=False)
            | filtered["Document Title"].fillna("").str.contains(term, case=False, regex=False)
            | filtered["Remarks"].fillna("").str.contains(term, case=False, regex=False)
        )
        filtered = filtered[mask]

    return filtered.sort_values(["Plan Issue Date", "Discipline", "Document No."], na_position="last")


def render_empty_state() -> None:
    st.warning("No records match the selected filters. Broaden or clear the filters.")


# Select the data source only from the same folder as app.py.
APP_DIR = Path(__file__).resolve().parent


@st.cache_data(show_spinner=False, max_entries=4)
def load_local_workbook(path_text: str, modified_ns: int, size_bytes: int) -> pd.DataFrame:
    # modified_ns and size_bytes are cache-key parts so file changes refresh the data.
    del modified_ns, size_bytes
    return read_document_master(Path(path_text))


def discover_local_workbook() -> tuple[Path, pd.DataFrame]:
    preferred = APP_DIR / SOURCE_FILENAME
    candidates = []
    if preferred.is_file():
        candidates.append(preferred)

    for candidate in sorted(APP_DIR.glob("*.xlsx")):
        if candidate.name.startswith("~$"):
            continue
        if candidate.name == "TKPR2_DML_filtered.xlsx":
            continue
        if candidate not in candidates:
            candidates.append(candidate)

    if not candidates:
        raise FileNotFoundError(
            "No .xlsx file was found in the same folder as app.py. "
            f"Recommended file name: {SOURCE_FILENAME}"
        )

    valid_sources: list[tuple[Path, pd.DataFrame]] = []
    failures: list[str] = []
    for candidate in candidates:
        try:
            stat = candidate.stat()
            frame = load_local_workbook(str(candidate), stat.st_mtime_ns, stat.st_size)
            if candidate == preferred:
                return candidate, frame
            valid_sources.append((candidate, frame))
        except Exception as exc:
            failures.append(f"{candidate.name}: {exc}")

    if len(valid_sources) == 1:
        return valid_sources[0]
    if len(valid_sources) > 1:
        # If multiple matching files exist, use the most recently modified one.
        return max(valid_sources, key=lambda item: item[0].stat().st_mtime_ns)

    detail = " | ".join(failures)
    raise WorkbookSchemaError(
        "No Excel file matching the TKPR2 template was found in the same folder. " + detail
    )


st.sidebar.header("Filters")
try:
    source_path, data = discover_local_workbook()
except Exception as exc:
    st.error(f"The Excel file could not be read: {exc}")
    st.info(
        "Keep app.py and the Excel file in the same folder. "
        f"Preferred file name: {SOURCE_FILENAME}"
    )
    st.stop()

min_plan_date = data["Plan Issue Date"].min()
max_plan_date = data["Plan Issue Date"].max()
if pd.isna(min_plan_date) or pd.isna(max_plan_date):
    st.error("No valid date was found in column M (Plan Issue date).")
    st.stop()

filter_keys = [
    "flt_disciplines",
    "flt_areas",
    "flt_responsible",
    "flt_types",
    "flt_actions",
    "flt_search",
]
discipline_options = normalised_options(data["Discipline"])
area_options = normalised_options(data["Area"])
responsible_options = normalised_options(data["Responsible"])
type_options = normalised_options(data["Deliverable Type"])
action_options = normalised_options(data["Document Action"])

if st.sidebar.button("Clear Filters", width="stretch"):
    for key in filter_keys:
        st.session_state.pop(key, None)
    st.rerun()

disciplines = st.sidebar.multiselect(
    "Discipline", discipline_options, default=[], placeholder="All", key="flt_disciplines"
)
areas = st.sidebar.multiselect(
    "Area", area_options, default=[], placeholder="All", key="flt_areas"
)
responsible = st.sidebar.multiselect(
    "Parner", responsible_options, default=[], placeholder="All", key="flt_responsible"
)
deliverable_types = st.sidebar.multiselect(
    "Deliverable Type", type_options, default=[], placeholder="All", key="flt_types"
)
actions = st.sidebar.multiselect(
    "New / Update", action_options, default=[], placeholder="All", key="flt_actions"
)
search_text = st.sidebar.text_input(
    "Search document no. / title / remarks",
    placeholder="e.g. PIP procedure",
    key="flt_search",
)

filtered = filter_frame(
    data,
    disciplines,
    areas,
    responsible,
    deliverable_types,
    actions,
    search_text,
)

title_col, refresh_col = st.columns([7.4, 1.25], gap="medium", vertical_alignment="center")
with title_col:
    st.title("TKPR2 Document Master List Dashboard")
with refresh_col:
    if st.button("REFRESH", type="primary", width="stretch"):
        st.cache_data.clear()
        st.rerun()
st.progress(0.05, text="Document progress for all documents: 5% — Publication Status: Not Published")

if filtered.empty:
    render_empty_state()
    st.stop()

placeholder_count = int(filtered["Is Placeholder"].sum())

# KPI cards navigate to a dedicated document-list view using the URL query parameter.
kpi_view = st.query_params.get("kpi")

kpi_filters = {
    "total": ("TOTAL DOCUMENT", filtered),
    "ret": ("RET SCOPE", filtered[filtered["Responsible"] == "RET"]),
    "toyo": ("TOYO SCOPE", filtered[filtered["Responsible"] == "TOYO"]),
    "external": ("External", filtered[filtered["Deliverable Type"] == "External"]),
    "internal": ("Internal", filtered[filtered["Deliverable Type"] == "Internal"]),
    "progress": ("Document Progress", filtered[filtered["Document Progress"] < 1]),
}

# Dedicated KPI detail page: only the selected document list is shown.
if kpi_view in kpi_filters:
    detail_title, detail_frame = kpi_filters[kpi_view]
    back_url = "?"
    st.markdown(
        f"<a href='{back_url}' target='_self' style='text-decoration:none;color:#1267A5;font-weight:800;font-size:16px;'>‹ BACK TO DASHBOARD</a>",
        unsafe_allow_html=True,
    )
    st.markdown(f"## {detail_title}")
    count_col, download_col = st.columns([5, 1.35], vertical_alignment="center")
    with count_col:
        st.caption(f"{fmt_int(len(detail_frame))} documents")
    with download_col:
        safe_name = detail_title.lower().replace(" ", "_").replace("/", "_")
        st.download_button(
            "EXCEL DOWNLOAD",
            data=build_export_xlsx(detail_frame.copy()),
            file_name=f"{safe_name}_document_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
            key=f"download_{kpi_view}",
        )
    render_document_list_table(detail_frame.copy())
    st.stop()

# Dashboard KPI cards. Clicking the card opens the dedicated list page in the same browser tab.
kpi_cards = [
    ("TOTAL DOCUMENT", fmt_int(len(filtered)), f"Total {fmt_int(len(data))}", "total"),
    ("RET SCOPE", fmt_int((filtered["Responsible"] == "RET").sum()), "RET documents", "ret"),
    ("TOYO SCOPE", fmt_int((filtered["Responsible"] == "TOYO").sum()), "TOYO documents", "toyo"),
    ("External", fmt_int((filtered["Deliverable Type"] == "External").sum()), "External documents", "external"),
    ("Internal", fmt_int((filtered["Deliverable Type"] == "Internal").sum()), "Internal documents", "internal"),
    ("Document Progress", "5%", "Not Published", "progress"),
]

kpi_cols = st.columns(6, gap="small")
for col, (label, value, note, key) in zip(kpi_cols, kpi_cards):
    with col:
        st.markdown(
            f"""
            <a href="?kpi={key}" target="_self" class="kpi-nav-card">
              <span class="kpi-nav-label">{label}</span>
              <span class="kpi-nav-value">{value}</span>
              <span class="kpi-nav-note">{note}</span>
            </a>
            """,
            unsafe_allow_html=True,
        )


summary_tab, schedule_tab, detail_tab = st.tabs(
    ["Overview", "Plan Schedule", "Document List"]
)

with summary_tab:
    discipline_col, ownership_col, action_col = st.columns([1.8, 1, 1], gap="medium")
    with discipline_col:
        discipline_counts = (
            filtered.groupby("Discipline", dropna=False)
            .size()
            .reset_index(name="Documents")
            .sort_values("Documents", ascending=False)
        )
        fig_discipline = px.bar(
            discipline_counts,
            x="Discipline",
            y="Documents",
            title="Document Count by Discipline",
            color="Documents",
            color_continuous_scale=["#BDE5E5", "#2F80C1", "#0B3A63"],
            text_auto=True,
            category_orders={"Discipline": discipline_counts["Discipline"].tolist()},
        )
        fig_discipline.update_traces(
            textposition="outside",
            cliponaxis=False,
            textfont=dict(size=11, color="#173A57", family="Segoe UI"),
        )
        fig_discipline.update_xaxes(
            tickangle=-45,
            title_text="Discipline",
            title_font=dict(size=13, color="#183B5B", family="Segoe UI"),
            tickfont=dict(size=10, color="#304B63", family="Segoe UI"),
            linecolor="#9FB3C5",
            tickcolor="#7F96AA",
            showline=True,
            showgrid=False,
            automargin=True,
        )
        fig_discipline.update_yaxes(
            title_text="Documents",
            title_font=dict(size=13, color="#183B5B", family="Segoe UI"),
            tickfont=dict(size=10, color="#304B63", family="Segoe UI"),
            rangemode="tozero",
            gridcolor="#E2E9F0",
            zerolinecolor="#CBD8E4",
            linecolor="#9FB3C5",
            tickcolor="#7F96AA",
            showline=True,
            automargin=True,
        )
        fig_discipline.update_layout(
            template="plotly_white",
            height=430,
            coloraxis_showscale=False,
            paper_bgcolor="#ffffff",
            plot_bgcolor="#ffffff",
            font=dict(color="#1B334A", size=11, family="Segoe UI"),
            title=dict(font=dict(size=16, color="#082E4F"), x=0.02, xanchor="left"),
            margin=dict(l=45, r=18, t=65, b=78),
        )
        st.plotly_chart(fig_discipline, width="stretch")

    with ownership_col:
        ownership = (
            filtered.groupby(["Responsible", "Deliverable Type"], dropna=False)
            .size()
            .reset_index(name="Documents")
        )
        # One stacked bar per partner: RET and TOYO have their own color families.
        fig_ownership = px.bar(
            ownership,
            x="Responsible",
            y="Documents",
            color="Deliverable Type",
            barmode="stack",
            title="Partner & Deliverable Type Distribution",
            text_auto=True,
            color_discrete_map={"External": "#2563EB", "Internal": "#93C5FD"},
            labels={"Responsible": "Partner"},
        )
        partner_colors = {
            ("RET", "External"): "#2563EB",
            ("RET", "Internal"): "#93C5FD",
            ("TOYO", "External"): "#F97316",
            ("TOYO", "Internal"): "#FDBA74",
        }
        for trace in fig_ownership.data:
            trace.marker.color = [
                partner_colors.get((str(x), str(trace.name)), "#94A3B8")
                for x in trace.x
            ]
            trace.marker.line.color = "#FFFFFF"
            trace.marker.line.width = 1.5

        for trace in fig_ownership.data:
            trace.texttemplate = [f"{trace.name}<br><b>{int(v)}</b>" for v in trace.y]

        fig_ownership.update_traces(
            textposition="inside",
            textfont=dict(size=11, color="#FFFFFF", family="Segoe UI", weight="bold"),
            insidetextanchor="middle",
            cliponaxis=False,
        )
        fig_ownership.update_xaxes(
            showgrid=False,
            title_text="Partner",
            title_font=dict(size=13, color="#183B5B", family="Segoe UI"),
            tickfont=dict(size=11, color="#304B63", family="Segoe UI"),
            linecolor="#9FB3C5",
            tickcolor="#7F96AA",
            showline=True,
            automargin=True,
        )
        fig_ownership.update_yaxes(
            title_text="Documents",
            title_font=dict(size=13, color="#183B5B", family="Segoe UI"),
            tickfont=dict(size=10, color="#304B63", family="Segoe UI"),
            rangemode="tozero",
            gridcolor="#E2E9F0",
            zerolinecolor="#CBD8E4",
            linecolor="#9FB3C5",
            tickcolor="#7F96AA",
            showline=True,
            automargin=True,
        )
        fig_ownership.update_layout(
            showlegend=False,
            template="plotly_white",
            height=430,
            paper_bgcolor="#ffffff",
            plot_bgcolor="#ffffff",
            font=dict(color="#1B334A", size=11, family="Segoe UI"),
            title=dict(font=dict(size=16, color="#082E4F"), x=0.02, xanchor="left"),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.01,
                xanchor="right",
                x=1,
                title=dict(text="", font=dict(size=12, color="#183B5B")),
                font=dict(size=11, color="#304B63", family="Segoe UI"),
                bgcolor="rgba(255,255,255,.96)",
            ),
            margin=dict(l=50, r=18, t=88, b=55),
        )
        st.plotly_chart(fig_ownership, width="stretch")

    with action_col:
        action_counts = filtered["Document Action"].value_counts().rename_axis("Action").reset_index(name="Documents")
        fig_action = px.pie(
            action_counts,
            names="Action",
            values="Documents",
            hole=.55,
            title="New / Update Distribution",
            color_discrete_map={"New": "#1E6FAE", "Update": "#E67E22"},
        )
        fig_action.update_traces(
            textposition="inside",
            textinfo="percent+label",
            textfont=dict(size=12, color="#FFFFFF", family="Segoe UI"),
            marker=dict(line=dict(color="#ffffff", width=2)),
        )
        fig_action.update_layout(
            template="plotly_white",
            height=430,
            paper_bgcolor="#ffffff",
            plot_bgcolor="#ffffff",
            font=dict(color="#1B334A", size=11, family="Segoe UI"),
            title=dict(font=dict(size=16, color="#082E4F"), x=0.02, xanchor="left"),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.04,
                xanchor="center",
                x=0.5,
                title=dict(text="", font=dict(size=12, color="#183B5B")),
                font=dict(size=11, color="#304B63", family="Segoe UI"),
                bgcolor="rgba(255,255,255,.96)",
            ),
            margin=dict(l=18, r=18, t=65, b=55),
        )
        st.plotly_chart(fig_action, width="stretch")

    st.subheader("Summary by Area")
    area_summary = (
        filtered.groupby("Area", dropna=False)
        .agg(
            Documents=("Document No.", "size"),
            RET=("Responsible", lambda s: int((s == "RET").sum())),
            TOYO=("Responsible", lambda s: int((s == "TOYO").sum())),
            External=("Deliverable Type", lambda s: int((s == "External").sum())),
            Internal=("Deliverable Type", lambda s: int((s == "Internal").sum())),
            Document_Progress=("Document Progress", "mean"),
        )
        .reset_index()
        .sort_values("Documents", ascending=False)
    )
    render_area_summary_table(area_summary)

with schedule_tab:
    monthly = (
        filtered.groupby(["Plan Month", "Responsible"], dropna=False)
        .size()
        .reset_index(name="Documents")
        .sort_values("Plan Month")
    )
    fig_monthly = px.bar(
        monthly,
        x="Plan Month",
        y="Documents",
        color="Responsible",
        barmode="stack",
        title="Monthly Plan by Plan Issue Date (Column M)",
        color_discrete_map={"RET": "#0B8F92", "TOYO": "#1E6FAE"},
        text_auto=True,
        labels={"Responsible": "Parner"},
    )
    fig_monthly.update_traces(
        textfont=dict(color="#FFFFFF", size=11, family="Segoe UI"),
        marker_line_color="#FFFFFF",
        marker_line_width=0.7,
    )
    fig_monthly.update_xaxes(
        dtick="M1",
        tickformat="%b\n%Y",
        title_text="Plan Month",
        title_font=dict(size=14, color="#183B5B", family="Segoe UI"),
        tickfont=dict(size=11, color="#304B63", family="Segoe UI"),
        linecolor="#9FB3C5",
        linewidth=1.2,
        tickcolor="#7F96AA",
        ticks="outside",
        showline=True,
        showgrid=False,
        automargin=True,
    )
    fig_monthly.update_yaxes(
        gridcolor="#E2E9F0",
        zerolinecolor="#CBD8E4",
        title_text="Documents",
        title_font=dict(size=14, color="#183B5B", family="Segoe UI"),
        tickfont=dict(size=11, color="#304B63", family="Segoe UI"),
        linecolor="#9FB3C5",
        linewidth=1.2,
        tickcolor="#7F96AA",
        ticks="outside",
        showline=True,
        automargin=True,
    )
    fig_monthly.update_layout(
        template="plotly_white",
        height=520,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#1B334A", size=12, family="Segoe UI"),
        title=dict(
            font=dict(size=18, color="#082E4F", family="Segoe UI"),
            x=0.02,
            xanchor="left",
        ),
        legend=dict(
            title=dict(text="Parner", font=dict(size=13, color="#183B5B", family="Segoe UI")),
            font=dict(size=12, color="#304B63", family="Segoe UI"),
            bgcolor="rgba(255,255,255,.96)",
            bordercolor="#D2DEE9",
            borderwidth=1,
            itemsizing="constant",
        ),
        margin=dict(l=58, r=28, t=72, b=55),
    )
    st.plotly_chart(fig_monthly, width="stretch")

    monthly_table = (
        filtered.assign(Plan_Month_Label=filtered["Plan Issue Date"].dt.strftime("%Y-%m"))
        .groupby(["Plan_Month_Label", "Responsible"], dropna=False)
        .size()
        .unstack(fill_value=0)
        .sort_index()
    )
    monthly_table["Total"] = monthly_table.sum(axis=1)
    total_row = monthly_table.sum(axis=0).to_frame().T
    total_row.index = ["Total"]
    monthly_table = pd.concat([monthly_table, total_row]).reset_index()
    # Compact, colorful and centered monthly summary table.
    monthly_view = monthly_table.copy()
    first_col = monthly_view.columns[0]
    monthly_view = monthly_view.rename(columns={first_col: "Month"})

    def monthly_row_colors(row):
        # Color by row instead of by column for a cleaner monthly view.
        if str(row.get("Month", "")) == "Total":
            return ["background-color:#16324F;color:#FFFFFF;font-weight:800" for _ in row]
        row_index = monthly_view.index.get_loc(row.name)
        color = "#F3F8FC" if row_index % 2 == 0 else "#E4F1F7"
        return [f"background-color:{color};color:#12344D" for _ in row]

    monthly_styler = (
        monthly_view.style
        .set_properties(**{
            "text-align": "center",
            "vertical-align": "middle",
            "font-weight": "600",
            "border": "1px solid #D4E0E9",
        })
        .set_table_styles([
            {"selector": "th", "props": [
                ("text-align", "center"),
                ("vertical-align", "middle"),
                ("font-weight", "800"),
                ("background-color", "#0F4C81"),
                ("color", "white"),
                ("border", "1px solid #CBD5E1"),
            ]},
            {"selector": "td", "props": [("text-align", "center")]},
        ])
        .apply(monthly_row_colors, axis=1)
    )

    table_left, table_center, table_right = st.columns([1.5, 7, 1.5])
    with table_center:
        st.dataframe(
            monthly_styler,
            hide_index=True,
            width="stretch",
            height=min(430, 38 + max(len(monthly_view), 1) * 35),
            column_config={
                "Month": st.column_config.TextColumn("Month", width="small"),
                "RET": st.column_config.NumberColumn("RET", width="small"),
                "TOYO": st.column_config.NumberColumn("TOYO", width="small"),
                "Total": st.column_config.NumberColumn("Total", width="small"),
            },
        )

with detail_tab:
    detail_frame = filtered.copy()
    detail_frame["Plan Issue Date"] = pd.to_datetime(
        detail_frame["Plan Issue Date"], errors="coerce"
    )
    detail_placeholder_count = int(detail_frame["Is Placeholder"].sum())

    st.caption(
        f"{fmt_int(len(detail_frame))} records displayed. "
        f"Document numbers containing an XXXX placeholder: {fmt_int(detail_placeholder_count)}. "
        "Search within each column and click column titles to sort, including Plan Issue Date."
    )
    download_spacer, excel_download_col = st.columns([8, 1], gap="small")
    with excel_download_col:
        st.download_button(
            "Excel Download",
            data=build_export_xlsx(detail_frame),
            file_name="Document_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    render_document_list_table(detail_frame)

