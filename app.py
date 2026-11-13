import os
import io
import re
import json
import base64
import html
import difflib
import urllib.parse
from datetime import datetime, date

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import openpyxl

st.set_page_config(
    page_title="TIU · PTM Analytics Portal",
    page_icon=":material/school:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
RECORDS_PATH = os.path.join(APP_DIR, "ptm_records.json")
DEFAULT_MARKS = os.path.join(APP_DIR, "PTM_BTECH_2025_DATA with attendance 13 Oct.xlsx")
DEFAULT_ATT = os.path.join(APP_DIR, "SOF 2ND YEAR 3RD SEM CLASS WISE ATTENDANCE (1) (10).xlsx")
LOGO_PATH = os.path.join(APP_DIR, "assets", "logo_header.png")
# Local mode (workbooks sit next to the app): PTM records persist to a file.
# Hosted mode (no bundled workbooks): records stay in the visitor's own session only,
# so one user's student notes are never visible to another.
PERSIST = os.path.exists(DEFAULT_MARKS)


def _logo_b64():
    try:
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""


LOGO_B64 = _logo_b64()

# ── Palette ───────────────────────────────────────────────────────────────────
NAVY, BLUE, SLATE, LIGHT = "#1e3a5f", "#2563eb", "#94a3b8", "#cbd5e1"
GREEN, AMBER, RED = "#059669", "#d97706", "#dc2626"
RISK_ORDER = ["Critical", "Needs Improvement", "On Track"]
RISK_COLORS = {"On Track": GREEN, "Needs Improvement": AMBER, "Critical": RED}
PTM_STATUSES = ["Not met", "Met", "Follow-up required", "Parent absent"]
VIEWS = ["Overview", "Student Dossier", "Directory", "Subjects", "Interventions"]
VIEW_LABELS = {"Overview": "Overview", "Student Dossier": "Student", "Directory": "Directory",
               "Subjects": "Subjects", "Interventions": "Actions"}
VIEW_ICONS = {
    "Overview": ":material/dashboard:",
    "Student Dossier": ":material/person:",
    "Directory": ":material/groups:",
    "Subjects": ":material/menu_book:",
    "Interventions": ":material/priority_high:",
}
esc = lambda v: html.escape(str(v))

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');

html, body, [data-testid="stAppViewContainer"], section.main, .stMarkdown, button, input, textarea, select {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}
[data-testid="stAppViewContainer"] { background: #ffffff; }
[data-testid="stHeader"] { background: rgba(255,255,255,0.9); backdrop-filter: blur(6px); }
[data-testid="stSidebar"] { background: #f8fafc; border-right: 1px solid #e2e8f0; }
.block-container { padding-top: 3.6rem !important; padding-bottom: 3rem !important; max-width: 1480px; }
#MainMenu, footer { visibility: hidden; }

h1, h2, h3 { font-family: 'Hanken Grotesk', sans-serif !important; letter-spacing: -0.015em; }

/* Header */
.portal-header { display:flex; justify-content:space-between; align-items:center; gap:16px;
    padding-bottom: 14px; margin-bottom: 6px; border-bottom: 1px solid #e2e8f0; flex-wrap: wrap; }
.brand { display:flex; align-items:center; gap:14px; }
.brand-logo { height:40px; width:auto; display:block; }
.brand-divider { width:1px; height:38px; background:#e2e8f0; }
.kpi[title] { cursor:help; }
.badge[title], .meta-chip[title] { cursor:help; }
.brand-mark { width:42px; height:42px; border-radius:10px; background:#1e3a5f; color:#fff;
    display:flex; align-items:center; justify-content:center; font-family:'Hanken Grotesk',sans-serif;
    font-weight:800; font-size:1rem; letter-spacing:0.02em; }
.portal-title { font-family:'Hanken Grotesk',sans-serif; font-size:1.45rem; font-weight:800; color:#0f172a;
    letter-spacing:-0.02em; margin:0; line-height:1.15; }
.portal-subtitle { font-size:0.78rem; color:#64748b; margin-top:3px; font-weight:500; }
.meta-row { display:flex; gap:8px; flex-wrap:wrap; }
.meta-chip { background:#f8fafc; border:1px solid #e2e8f0; border-radius:999px; padding:5px 12px;
    font-size:0.76rem; color:#475569; font-weight:500; white-space:nowrap; }
.meta-chip b { color:#0f172a; }

/* Section headings */
.sh { font-family:'Hanken Grotesk',sans-serif; font-size:0.95rem; font-weight:700; color:#0f172a;
    margin: 6px 0 2px; }
.sh-sub { font-size:0.78rem; color:#64748b; margin-bottom:10px; }

/* Profile card */
.profile-card { background: linear-gradient(135deg, #1e3a5f 0%, #24476f 100%); border-radius:12px;
    padding:22px 26px; display:flex; align-items:center; gap:20px; color:#fff; margin-bottom:14px;
    flex-wrap: wrap; }
.profile-initials { width:58px; height:58px; border-radius:50%; background:rgba(255,255,255,0.14);
    border:1px solid rgba(255,255,255,0.25); display:flex; align-items:center; justify-content:center;
    font-family:'Hanken Grotesk',sans-serif; font-size:1.3rem; font-weight:700; flex-shrink:0; }
.profile-name { font-family:'Hanken Grotesk',sans-serif; font-size:1.4rem; font-weight:700; margin:0;
    letter-spacing:-0.01em; }
.profile-meta { display:flex; flex-wrap:wrap; gap:6px 22px; font-size:0.8rem; color:#cbd5e1; margin-top:6px; }
.profile-meta b { color:#ffffff; font-weight:600; }
.profile-badges { display:flex; flex-direction:column; gap:6px; align-items:flex-end; margin-left:auto; }

/* Badges */
.badge { display:inline-block; padding:3px 10px; border-radius:999px; font-size:0.72rem; font-weight:600;
    letter-spacing:0.01em; white-space:nowrap; }
.badge-safe   { background:#ecfdf5; color:#065f46; border:1px solid #a7f3d0; }
.badge-warn   { background:#fffbeb; color:#92400e; border:1px solid #fde68a; }
.badge-danger { background:#fef2f2; color:#991b1b; border:1px solid #fecaca; }
.badge-muted  { background:#f1f5f9; color:#475569; border:1px solid #e2e8f0; }
.badge-info   { background:#eff6ff; color:#1e40af; border:1px solid #bfdbfe; }

/* KPI cards */
.kpi-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:12px; margin-bottom:16px; }
.kpi { background:#fff; border:1px solid #e2e8f0; border-radius:10px; padding:14px 16px; position:relative; overflow:hidden; }
.kpi::before { content:""; position:absolute; left:0; top:0; bottom:0; width:3px; background: var(--accent, #1e3a5f); }
.kpi-label { font-size:0.7rem; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.05em; }
.kpi-value { font-family:'Hanken Grotesk',sans-serif; font-size:1.7rem; font-weight:700; color:#0f172a;
    margin:6px 0 4px; line-height:1; letter-spacing:-0.02em; font-variant-numeric: tabular-nums; }
.kpi-sub { font-size:0.74rem; color:#64748b; }
.kpi-bar { height:4px; border-radius:2px; background:#f1f5f9; margin-top:9px; overflow:hidden; }
.kpi-bar-fill { height:100%; border-radius:2px; }

/* Insight panel */
.insight-panel { background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:14px 18px; height:100%; }
.insight-panel h4 { font-family:'Hanken Grotesk',sans-serif; font-size:0.8rem; text-transform:uppercase;
    letter-spacing:0.05em; color:#475569; margin:0 0 10px; }
.insight-item { display:flex; gap:10px; font-size:0.85rem; color:#0f172a; line-height:1.5; margin-bottom:8px; }
.insight-item:last-child { margin-bottom:0; }
.dot { width:8px; height:8px; border-radius:50%; margin-top:7px; flex-shrink:0; }

/* Tables */
.ptm-table { width:100%; border-collapse:separate; border-spacing:0; font-size:0.82rem; background:#fff;
    border:1px solid #e2e8f0; border-radius:8px; overflow:hidden; }
.ptm-table th { background:#f8fafc; color:#475569; font-weight:600; text-transform:uppercase; font-size:0.68rem;
    letter-spacing:0.05em; padding:9px 12px; border-bottom:1px solid #e2e8f0; text-align:left; }
.ptm-table td { padding:8px 12px; border-bottom:1px solid #f1f5f9; color:#0f172a; font-variant-numeric:tabular-nums; }
.ptm-table tr:last-child td { border-bottom:none; }
.ptm-table tr:hover td { background:#f8fafc; }
.text-right { text-align:right !important; } .text-center { text-align:center !important; }

/* Simulator */
.sim { border-radius:10px; padding:14px 16px; font-size:0.86rem; line-height:1.55; }
.sim-ok    { background:#ecfdf5; border:1px solid #a7f3d0; color:#065f46; }
.sim-alert { background:#fef2f2; border:1px solid #fecaca; color:#991b1b; }
.sim-num   { font-family:'Hanken Grotesk',sans-serif; font-size:1.6rem; font-weight:700; display:block; margin:2px 0; }

.empty-state { text-align:center; padding:48px 16px; color:#64748b; border:1px dashed #cbd5e1; border-radius:12px; background:#f8fafc; }
.empty-state b { color:#0f172a; display:block; font-size:1rem; margin-bottom:4px; }

/* Widgets */
[data-testid="stVerticalBlockBorderWrapper"] { border-radius:10px; }
div[data-testid="stSegmentedControl"] button { font-weight:600; }
[data-testid="stMetricValue"] { font-family:'Hanken Grotesk',sans-serif; }


/* Student cards (used by the mobile layout) */
.scard { padding:0; margin:0; }
.scard-top { display:flex; justify-content:space-between; align-items:flex-start; gap:10px; }
.scard-name { font-family:'Hanken Grotesk',sans-serif; font-weight:700; font-size:1rem; color:#0f172a; line-height:1.25; }
.scard-sub { font-size:0.74rem; color:#64748b; margin-top:2px; }
.scard-stats { display:grid; grid-template-columns:repeat(3,1fr); gap:8px; margin-top:10px; }
.scard-stats span { display:block; font-size:0.62rem; text-transform:uppercase; letter-spacing:.05em; color:#64748b; font-weight:600; }
.scard-stats b { font-family:'Hanken Grotesk',sans-serif; font-size:1.05rem; }
.scard-extra { margin:8px 0 12px; }
[class*="st-key-cb_"] { margin-bottom:4px; }
.st-key-cards_dir, .st-key-cards_int { display:none; }

@media (max-width: 760px) {
    .block-container { padding: 2.3rem 0.85rem 6.5rem !important; }
    [data-testid="stToolbar"] { display:none; }
    .portal-header { gap:8px; padding-bottom:10px; }
    .brand { gap:10px; }
    .brand-logo { height:26px; } .brand-divider { height:26px; }
    .portal-title { font-size:1.02rem; }
    .portal-subtitle, .meta-row { display:none; }
    h2 { font-size:1.2rem !important; }
    .kpi-grid { grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; }
    .kpi { padding:11px 12px; } .kpi-value { font-size:1.35rem; } .kpi-sub { font-size:0.7rem; }
    .profile-card { padding:16px; gap:12px; }
    .profile-initials { width:46px; height:46px; font-size:1.05rem; }
    .profile-name { font-size:1.15rem; }
    .profile-badges { flex-direction:row; flex-wrap:wrap; margin-left:0; width:100%; align-items:flex-start; }
    .profile-meta { gap:4px 14px; }
    .insight-panel { padding:12px 14px; }

    /* swap wide tables for card lists */
    .st-key-tbl_dir, .st-key-tbl_int { display:none; }
    .st-key-cards_dir, .st-key-cards_int { display:block; }

    /* bottom tab bar */
    .st-key-view { position:fixed !important; left:0; right:0; bottom:0; width:100% !important; z-index:1000;
        background:#ffffff; border-top:1px solid #e2e8f0; box-shadow:0 -4px 16px rgba(15,23,42,.08);
        padding:6px 6px calc(6px + env(safe-area-inset-bottom)); }
    .st-key-view [data-testid="stButtonGroup"] { width:100% !important; }
    .st-key-view [data-baseweb="button-group"] { display:flex !important; flex-wrap:nowrap !important; width:100% !important; max-width:none !important; gap:2px; }
    .st-key-view button { flex:1 1 0% !important; width:0 !important; max-width:none !important; min-width:0 !important; min-height:52px; border:none !important;
        background:transparent !important; padding:4px 2px !important; border-radius:10px !important;
        color:#64748b; }
    .st-key-view button > div, .st-key-view button span[data-has-shortcut] {
        flex-direction:column !important; align-items:center; justify-content:center; gap:2px !important; }
    .st-key-view button p { font-size:0.6rem; letter-spacing:-0.01em; font-weight:600; margin:0; line-height:1.1; white-space:nowrap; }
    .st-key-view button[kind="segmented_controlActive"] { color:#1e3a5f; background:#eaf0f9 !important; }

    /* filters: two per row */
    [data-testid="stExpander"] [data-testid="stHorizontalBlock"] { flex-wrap:wrap !important; gap:8px !important; }
    [data-testid="stExpander"] [data-testid="stColumn"] { flex:1 1 calc(50% - 8px) !important;
        min-width:calc(50% - 8px) !important; width:auto !important; }
    /* previous / next side by side, report full width */
    .st-key-dnav [data-testid="stHorizontalBlock"] { flex-wrap:wrap !important; gap:8px !important; }
    .st-key-dnav [data-testid="stColumn"] { flex:1 1 100% !important; min-width:100% !important; }
    .st-key-dnav [data-testid="stColumn"]:nth-child(-n+2) { flex:1 1 calc(50% - 8px) !important; min-width:calc(50% - 8px) !important; }
    /* action rows never stack */
    .st-key-qa_row [data-testid="stHorizontalBlock"], [class*="st-key-cb_"] [data-testid="stHorizontalBlock"] {
        flex-wrap:nowrap !important; gap:8px !important; }
    .st-key-qa_row [data-testid="stColumn"], [class*="st-key-cb_"] [data-testid="stColumn"] {
        min-width:0 !important; flex:1 1 0 !important; }

    /* comfortable touch targets */
    .stButton button, .stDownloadButton button, [data-testid="stLinkButton"] a, [data-testid="stBaseLinkButton-secondary"],
    [data-testid="stBaseLinkButton-primary"] { min-height:46px; }
    [data-baseweb="select"] > div, .stTextInput input, .stDateInput input, .stNumberInput input { min-height:46px; font-size:16px; }
    .stTextArea textarea { font-size:16px; }
}

@media print {
    [data-testid="stSidebar"], [data-testid="stHeader"], button, .stButton, [data-testid="stSegmentedControl"] { display:none !important; }
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HOVER TOOLTIPS — one-line purpose for every control (CSS-only, keyed by widget key)
# ══════════════════════════════════════════════════════════════════════════════
# value: str -> whole widget; list -> one tip per option (segmented control / pills)
TIPS = {
    "view": ["Portfolio snapshot: KPIs, risk by section and cohort benchmarks.",
             "One student's full PTM profile, talking points and meeting record.",
             "Search, sort and browse every student in the current filter.",
             "Subject-wise pass rates, score spread and weak spots.",
             "Prioritised list of students who need parent or faculty action."],
    "finder": ("Jump straight to any student by typing a name or roll number.", "right"),
    "f_stream": "Narrow every view to one programme stream (CC, DS, AIML ...).",
    "f_section": "Narrow every view to a single class section.",
    "f_att": "Filter by attendance band relative to the policy requirement.",
    "f_perf": "Filter by aggregate marks: high achievers, average or below pass.",
    "f_risk": "Show only students in one risk band.",
    "f_ptm": "Track outreach: has the parent meeting happened yet?",
    "btn_reset": "Clear all filters and show every student again.",
    "btn_prev": "Open the previous student in the filtered list.",
    "btn_next": "Open the next student in the filtered list.",
    "btn_report": ("Download a one-page printable PTM report with the university logo.", "right"),
    "dir_q": "Type a name, roll number or parent phone; partial words and typos are fine.",
    "dir_sort": ["Alphabetical by student name.", "Ascending roll number.", "Lowest attendance first.",
                 "Highest marks first.", "Lowest marks first.", "Most critical risk band first."],
    "dl_dir": "Download the list currently shown as a CSV file.",
    "subj_pick": "Pick a subject to list every student scoring below the pass mark.",
    "int_kind": ["Everyone below the attendance or pass-mark threshold.", "Students under the attendance requirement.",
                 "Students under the pass mark.", "Students failing both attendance and marks."],
    "dl_xlsx": "Excel workbook with one sheet per section, ready for mentors.",
    "dl_csv": "Plain CSV of the flagged students.",
    "u_marks": "Upload the marks workbook (needs ID and Name columns).",
    "u_att": "Upload the class-wise attendance workbook (one sheet per section).",
    "term_year": "Year used for attendance dates written without one, e.g. 20th July.",
    "att_thr_s": "Minimum attendance a student needs to stay exam-eligible.",
    "pass_thr_s": "Minimum aggregate percentage to count as passing.",
    "exp_log": "Download every saved PTM record as a CSV log.",
    "sim_*": "Pick a target attendance % to see how many lectures are needed.",
    "ptm_status_*": "Outcome of the parent meeting for this student.",
    "ptm_att_*": "Who from the family attended the meeting.",
    "ptm_date_*": "Date the meeting took place.",
    "ptm_fup_*": "Optional date to review this student's progress again.",
    "ptm_remarks_*": "Notes and agreed actions; included in the parent message and report.",
    "btn_save_ptm": "Save this meeting record to ptm_records.json.",
}


def _q(t):
    return t.replace("\\", "").replace('"', "'")


def build_tip_css():
    base = ("position:absolute;z-index:999999;width:max-content;max-width:min(300px,80vw);"
            "background:#0f172a;color:#fff;font:500 .72rem/1.4 Inter,sans-serif;padding:6px 10px;"
            "border-radius:6px;pointer-events:none;opacity:0;transition:opacity .15s ease .4s;"
            "box-shadow:0 4px 12px rgba(15,23,42,.25);white-space:normal;text-align:left;")
    out = []
    for key, tip in TIPS.items():
        sel = f'[class*="st-key-{key[:-1]}"]' if key.endswith("*") else f'[class~="st-key-{key}"]'
        if isinstance(tip, list):
            # one tooltip per option, anchored directly under the option being hovered
            out.append(f"{sel} button {{position:relative;overflow:visible !important;}}")
            out.append(f"{sel} button::after {{content:'';{base}top:calc(100% + 6px);left:0;}}")
            for i, t in enumerate(tip, 1):
                out.append(f'{sel} button:nth-of-type({i}):hover::after {{content:"{_q(t)}";opacity:1;}}')
            out.append(f"{sel} button:active::after {{opacity:0 !important;}}")
        else:
            text, side = (tip, "left") if isinstance(tip, str) else tip
            pos = "right:0;" if side == "right" else "left:0;"
            out.append(f"{sel} {{position:relative;}}")
            out.append(f'{sel}::after {{content:"{_q(text)}";{base}top:calc(100% + 2px);{pos}}}')
            out.append(f"{sel}:hover::after {{opacity:1;}}")
            out.append(f"{sel}:focus-within::after, {sel}:active::after {{opacity:0 !important;}}")
    return "<style>@media (hover:hover) {" + "\n".join(out) + "}</style>"


st.markdown(build_tip_css(), unsafe_allow_html=True)

KPI_TIPS = {
    "Students in scope": "Students matching the current filters.",
    "Mean attendance": "Average attendance % (from the attendance workbook) across students in scope.",
    "Mean aggregate": "Average total marks as a percentage of the maximum.",
    "Below attendance rule": "Students under the attendance requirement.",
    "Below pass mark": "Students whose aggregate is under the pass mark.",
    "PTM consultations": "Parent meetings completed or awaiting follow-up, out of students in scope.",
    "Attendance": "Attendance % counted from the attendance workbook up to the cut-off date.",
    "Aggregate score": "Total marks across all subjects.",
    "Section rank": "Rank by total marks within the section (students with marks only).",
    "Cohort percentile": "Share of the cohort this student scored higher than.",
    "Subjects below pass": "Number of subjects scored under the pass mark.",
    "Subjects": "Subjects found in the marks sheet.",
    "Lowest pass rate": "The subject where the fewest students passed.",
    "Highest pass rate": "The subject where the most students passed.",
    "Failing grades": "Total subject results under the pass mark.",
    "Flagged students": "Students under the attendance requirement or the pass mark.",
    "Critical": "Students in the highest-risk band.",
    "Avg lectures to recover": "Average consecutive lectures needed to reach the requirement.",
    "Not yet consulted": "Flagged students whose PTM status is still Not met.",
    "With phone on file": "Flagged students with a parent number available.",
}


# ══════════════════════════════════════════════════════════════════════════════
# DATA ENGINE
# ══════════════════════════════════════════════════════════════════════════════
MONTHS = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
          'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
TERM_MONTHS = ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

GENERIC_LABELS = {
    'session', 'elective', 'elective/session', 'general session',
    'no faculty', 'no class', 'tba', 'n/a', 'none', '', 'class reschedule',
    'group a', 'group b', 'grp a', 'grp b'
}


def parse_date(val, year=None):
    if isinstance(val, datetime):
        return val
    if not val:
        return None
    s = str(val).strip()
    m = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)', s)
    if m:
        mo = m.group(2)[:3].lower()
        if mo in MONTHS:
            try:
                return datetime(year or datetime.now().year, MONTHS[mo], int(m.group(1)))
            except ValueError:
                pass
    if not re.search(r'\d', s):
        return None
    try:
        return pd.to_datetime(s).to_pydatetime()
    except Exception:
        return None


def clean_id(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    t = str(v).replace('\xa0', ' ').strip()
    m = re.search(r'\d{6,}', t)           # roll numbers are long digit runs; ignores "(lateral)" etc.
    return m.group(0) if m else re.sub(r'\.0$', '', t)


def clean_phone(v):
    if v is None:
        return ""
    d = re.sub(r'[^0-9]', '', re.sub(r'\.0$', '', str(v).strip()))
    return d[-10:] if len(d) >= 10 else ""


def subject_key(name):
    """Normalise subject titles so 'Problem solving using C' == 'Problem Solving using C'."""
    s = name.lower().replace('&', ' and ')
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()


def pretty_subject(name):
    return name.title() if name.isupper() or name.islower() else name


@st.cache_data(show_spinner="Reading student records…")
def load_raw(marks_bytes, att_bytes, term_year):
    M = lambda: io.BytesIO(marks_bytes)

    # ── 1. Marks sheet ────────────────────────────────────────────────────────
    xl = pd.ExcelFile(M())
    msheet = next((s for s in xl.sheet_names
                   if any(k in s.lower() for k in ['subject', 'marks', 'score', 'result'])),
                  xl.sheet_names[0])

    raw = pd.read_excel(M(), sheet_name=msheet, header=None, nrows=10)
    hrow = 0
    for i, row in raw.iterrows():
        rs = ' '.join(str(v).lower() for v in row if pd.notna(v))
        if any(k in rs for k in ['id', 'roll', 'reg']) and any(k in rs for k in ['name', 'student']):
            hrow = i
            break
    df_m = pd.read_excel(M(), sheet_name=msheet, skiprows=hrow)

    max_marks = None
    ws_tmp = openpyxl.load_workbook(M(), data_only=True)[msheet]
    for r in range(1, min(hrow + 2, ws_tmp.max_row + 1)):
        for c in range(1, ws_tmp.max_column + 1):
            found = re.search(r'out\s+of\s+(\d+)', str(ws_tmp.cell(r, c).value or ''), re.I)
            if found:
                max_marks = float(found.group(1))
                break
        if max_marks:
            break

    id_c = name_c = stream_c = total_c = att_c = None
    for col in df_m.columns:
        c = str(col).lower().strip()
        if re.search(r'\b(id|roll|reg|enrollment)\b', c) and not id_c:
            id_c = col
        elif re.search(r'\b(name|student)\b', c) and not name_c:
            name_c = col
        elif re.search(r'\b(stream|branch|dept|program|course)\b', c) and not stream_c:
            stream_c = col
        elif re.search(r'\b(grand\s*total|total\s*marks|aggregate)\b', c) and not total_c:
            total_c = col
        elif re.search(r'\b(attendance|att\s*%|att%)\b', c) and not att_c:
            att_c = col
    if not total_c:
        total_c = next((col for col in df_m.columns if 'total' in str(col).lower()), None)

    meta_set = {id_c, name_c, stream_c, total_c, att_c}
    subj_cols = [col for col in df_m.columns
                 if col not in meta_set
                 and pd.to_numeric(df_m[col], errors='coerce').notna().sum() >= min(5, len(df_m))]

    df_m['ID_clean'] = df_m[id_c].apply(clean_id) if id_c else ""
    df_m['NAME_clean'] = df_m[name_c].astype(str).str.strip() if name_c else "Unknown"
    df_m['STR_clean'] = df_m[stream_c].astype(str).str.strip() if stream_c else "General"
    for sc in subj_cols:
        df_m[sc] = pd.to_numeric(df_m[sc], errors='coerce')
    if total_c:
        df_m['Total_clean'] = pd.to_numeric(df_m[total_c], errors='coerce')
    else:
        df_m['Total_clean'] = df_m[subj_cols].sum(axis=1, min_count=1)
    # a blank total with no subject marks means "no marks on record", not zero
    df_m.loc[df_m[subj_cols].notna().sum(axis=1) == 0, 'Total_clean'] = np.nan
    df_m = df_m[df_m['ID_clean'].astype(str).str.len() > 0].reset_index(drop=True)

    # attendance cut-off date, e.g. column "Attendance % till 13 Oct"
    att_cutoff = None
    if att_c:
        mc = re.search(r'(?:till|upto|up to|as on|as of)\s+(.+)$', str(att_c), re.I)
        if mc:
            att_cutoff = parse_date(mc.group(1), term_year)

    if not max_marks:
        mx = df_m['Total_clean'].max()
        mx = 0 if pd.isna(mx) else mx
        max_marks = 1000. if mx > 500 else (float(int(np.ceil(mx / 100) * 100)) if mx > 100 else 100.)
    subj_max = max_marks / len(subj_cols) if subj_cols else 100.

    df_detail = None
    det_s = next((s for s in xl.sheet_names
                  if any(k in s.lower() for k in ['data', 'component', 'detail'])), None)
    if det_s and det_s != msheet:
        try:
            df_detail = pd.read_excel(M(), sheet_name=det_s)
            for c in df_detail.columns:
                if any(k in str(c).lower() for k in ['id', 'reg', 'roll']):
                    df_detail['ID_clean'] = df_detail[c].apply(clean_id)
                    break
            if 'ID_clean' not in df_detail.columns:
                df_detail = None
        except Exception:
            df_detail = None

    # ── 2. Attendance workbook ────────────────────────────────────────────────
    att_map, orphans = {}, []
    marks_ids = set(df_m['ID_clean'])
    if att_bytes:
        wb_att = openpyxl.load_workbook(io.BytesIO(att_bytes), data_only=True)
        for sname in wb_att.sheetnames:
            ws = wb_att[sname]
            hrow_a = aid = aname = acont = None
            for r in range(1, min(ws.max_row + 1, 6)):
                for c in range(1, min(ws.max_column + 1, 15)):
                    v = str(ws.cell(r, c).value or '').lower()
                    if re.search(r'\b(id|roll|reg)\b', v) and not aid:
                        aid = c
                    elif re.search(r'\b(name|student)\b', v) and not aname:
                        aname = c
                    elif re.search(r'\b(contact|phone|mobile)\b', v) and not acont:
                        acont = c
                if aid and aname:
                    hrow_a = r
                    break
            if not hrow_a:
                continue

            date_row = max(1, hrow_a - 1)
            col_dates, cur = [], None
            for c in range(1, ws.max_column + 1):
                d = parse_date(ws.cell(date_row, c).value, term_year)
                if d:
                    cur = d
                col_dates.append(cur)

            start_c = max(aid or 1, aname or 1, acont or 1) + 1

            # Pre-compute cleaned subject title per column (same for every row)
            col_sub = {}
            for c in range(start_c, ws.max_column + 1):
                raw_sub = str(ws.cell(hrow_a, c).value or '').strip()
                s = re.sub(r'\(.*?\)', '', raw_sub)
                s = re.sub(r'\d{1,2}[:.]\d{2}\s*(?:am|pm)?\s*-\s*\d{1,2}[:.]\d{2}\s*(?:am|pm)?', '', s, flags=re.I)
                s = re.sub(r'\s+', ' ', s).strip()
                col_sub[c] = None if (not s or s.lower() in GENERIC_LABELS or len(s) < 3) else s

            # read every student row once
            body = []
            for r in range(hrow_a + 1, ws.max_row + 1):
                sid = clean_id(ws.cell(r, aid).value)
                if not sid:
                    continue
                body.append((sid, clean_phone(ws.cell(r, acont).value) if acont else "",
                             {c: ws.cell(r, c).value for c in col_sub}))
                if sid not in marks_ids:
                    orphans.append((sname, sid))

            def state(val):
                """'P' present, 'A' absent, None = not applicable to this student (No class, GROUP A ...)."""
                sv = str(val).strip().lower() if val is not None else ''
                if val in (1, 1.0) or sv in ('1', 'p', 'present'):
                    return 'P'
                if val in (None, '', 0, 0.0) or sv in ('0', 'a', 'absent'):
                    return 'A'
                return None

            # a column is a real lecture only if it is within the cut-off AND at least one student was present
            held_cols = set()
            for c in col_sub:
                dt = col_dates[c - 1] if c - 1 < len(col_dates) else None
                if att_cutoff and dt and dt.date() > att_cutoff.date():
                    continue
                if any(state(v[c]) == 'P' for _, _, v in body):
                    held_cols.add(c)

            for sid, phone, vals in body:
                att_n = tot_n = 0
                day_s = {d: [0, 0] for d in WEEKDAYS}
                mon_s = {m: [0, 0] for m in TERM_MONTHS}
                sub_a, sub_disp = {}, {}
                for c in held_cols:
                    st_ = state(vals[c])
                    if st_ is None:
                        continue
                    is_p = st_ == 'P'
                    sub_nm = col_sub.get(c)
                    dt = col_dates[c - 1] if c - 1 < len(col_dates) else None
                    tot_n += 1
                    att_n += int(is_p)
                    if sub_nm:
                        k = subject_key(sub_nm)
                        sub_disp.setdefault(k, pretty_subject(sub_nm))
                        sub_a.setdefault(k, [0, 0])
                        sub_a[k][1] += 1
                        sub_a[k][0] += int(is_p)
                    if dt:
                        w, m2 = dt.strftime('%A'), dt.strftime('%b')
                        if w in day_s:
                            day_s[w][1] += 1
                            day_s[w][0] += int(is_p)
                        if m2 in mon_s:
                            mon_s[m2][1] += 1
                            mon_s[m2][0] += int(is_p)

                dropped = any(isinstance(v, str) and 'drop' in v.lower() for v in vals.values())
                att_map[sid] = dict(section=sname, contact=phone, attended=att_n, total=tot_n, dropped=dropped,
                                    pct=round(att_n / tot_n * 100, 2) if tot_n else None,
                                    day_s=day_s, mon_s=mon_s,
                                    sub_a={sub_disp[k]: v for k, v in sub_a.items()})

    # ── 3. Consolidate master records ─────────────────────────────────────────
    rows = []
    for _, row in df_m.iterrows():
        sid = str(row['ID_clean'])
        if not sid or sid.lower() == 'nan':
            continue
        ai = att_map.get(sid, {})
        official = float(row[att_c]) if att_c and pd.notna(row.get(att_c)) else None
        calc = ai.get('pct')
        final_att = calc if calc is not None else official      # attendance workbook is the source of truth
        total_m = row['Total_clean']
        has_marks = pd.notna(total_m)
        subj = {sc: float(row[sc]) for sc in subj_cols if pd.notna(row[sc])}
        rows.append(dict(
            ID=sid, Name=row['NAME_clean'], Stream=row['STR_clean'],
            Section=ai.get('section', f"{row['STR_clean']} (unassigned)"),
            Contact=ai.get('contact') or 'N/A',
            Grand_Total=float(total_m) if has_marks else np.nan, Max_Marks=max_marks,
            Marks_Pct=round(float(total_m) / max_marks * 100, 2) if has_marks else np.nan,
            Att_Pct=round(final_att, 2) if final_att is not None else np.nan,
            Att_Calc=calc if calc is not None else np.nan,
            Att_Official=official if official is not None else np.nan,
            In_Att_Sheet=bool(ai), Dropped_Out=bool(ai.get('dropped', False)),
            Attended=ai.get('attended'), Total_Lec=ai.get('total'),
            Day_S=ai.get('day_s', {}), Mon_S=ai.get('mon_s', {}), Sub_A=ai.get('sub_a', {}),
            Subj=subj,
        ))

    mdf = pd.DataFrame(rows)
    if not mdf.empty:
        mdf['Batch_Rank'] = mdf['Grand_Total'].rank(ascending=False, method='min')
        mdf['Section_Rank'] = mdf.groupby('Section')['Grand_Total'].rank(ascending=False, method='min')
        mdf['Section_Size'] = mdf.groupby('Section')['Grand_Total'].transform('count')   # students with marks
        mdf['Batch_Size'] = int(mdf['Grand_Total'].notna().sum())
        mdf['Percentile'] = (mdf['Grand_Total'].rank(pct=True) * 100).round(0)

    meta = dict(subj_cols=subj_cols, max_marks=max_marks, subj_max=subj_max,
                has_att=bool(att_map), has_detail=df_detail is not None,
                att_cutoff=att_cutoff.strftime('%d %b %Y') if att_cutoff else None,
                att_col=str(att_c) if att_c else None, att_orphans=len(orphans),
                att_sheets=len(set(a for a, _ in orphans)))
    return mdf, df_m, df_detail, meta


def classify(base, att_thr, pass_thr, subj_max):
    """Threshold-dependent fields — cheap, recomputed when sliders move."""
    mdf = base.copy()
    if mdf.empty:
        return mdf
    att = mdf['Att_Pct'].fillna(100.0)  # unknown attendance doesn't trigger an attendance flag
    crit = (att < att_thr - 15) | (mdf['Marks_Pct'] < pass_thr)
    warn = (att < att_thr) | (mdf['Marks_Pct'] < pass_thr + 15)
    mdf['Risk'] = np.where(crit, 'Critical', np.where(warn, 'Needs Improvement', 'On Track'))
    mdf['Badge_Cls'] = mdf['Risk'].map({'Critical': 'badge-danger', 'Needs Improvement': 'badge-warn',
                                        'On Track': 'badge-safe'})
    cut = pass_thr / 100 * subj_max
    mdf['Failed_Subj'] = mdf['Subj'].apply(lambda d: sum(v < cut for v in d.values()))
    tl = pd.to_numeric(mdf['Total_Lec'], errors='coerce')
    # lectures attended consistent with the displayed percentage (so the recovery maths matches)
    eff = (mdf['Att_Pct'] / 100 * tl).round()
    mdf['Attended_Eff'] = eff.where(mdf['Att_Pct'].notna() & tl.notna(), pd.to_numeric(mdf['Attended'], errors='coerce'))
    mdf['Lec_Needed'] = [lectures_needed(a, t, att_thr) for a, t in zip(mdf['Attended_Eff'], tl)]
    return mdf


def lectures_needed(attended, total, goal):
    if attended is None or total is None or pd.isna(attended) or pd.isna(total) or goal >= 100:
        return 0
    need = goal * total - 100 * attended
    return int(np.ceil(need / (100 - goal))) if need > 0 else 0


def att_band(pct, thr):
    if pd.isna(pct):
        return 'badge-muted', 'No data'
    if pct >= thr:
        return 'badge-safe', 'Eligible'
    if pct >= thr - 15:
        return 'badge-warn', 'Borderline'
    return 'badge-danger', 'Critical'


# ══════════════════════════════════════════════════════════════════════════════
# PTM RECORD STORE (persisted to ptm_records.json next to the app)
# ══════════════════════════════════════════════════════════════════════════════
def load_records():
    if 'ptm_records' not in st.session_state:
        st.session_state.ptm_records = {}
        if PERSIST:
            try:
                with open(RECORDS_PATH, 'r', encoding='utf-8') as f:
                    st.session_state.ptm_records = json.load(f)
            except Exception:
                pass
    return st.session_state.ptm_records


def save_record(sid, rec):
    recs = load_records()
    rec['updated'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    recs[sid] = rec
    if not PERSIST:
        return True
    try:
        with open(RECORDS_PATH, 'w', encoding='utf-8') as f:
            json.dump(recs, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════════════════
# UI HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def style_fig(fig, height=300, legend=False, static=True):
    fig.update_layout(
        height=height, margin=dict(l=4, r=4, t=8, b=4),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, system-ui, sans-serif", size=11, color="#334155"),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title=None),
        hoverlabel=dict(bgcolor="#0f172a", font_color="#ffffff", font_family="Inter", bordercolor="#0f172a"),
        bargap=0.25,
    )
    if static:                      # keeps one-finger scrolling working over charts on touch screens
        fig.update_layout(dragmode=False)
    fig.update_xaxes(gridcolor="#f1f5f9", zeroline=False, linecolor="#e2e8f0", automargin=True)
    fig.update_yaxes(gridcolor="#f1f5f9", zeroline=False, linecolor="#e2e8f0", automargin=True)
    return fig


CHART_CFG = {"displayModeBar": False}


def show_fig(fig):
    st.plotly_chart(fig, width="stretch", config=CHART_CFG, theme=None)


def section_heading(title, sub=None):
    st.markdown(f'<div class="sh">{esc(title)}</div>' +
                (f'<div class="sh-sub">{esc(sub)}</div>' if sub else ''), unsafe_allow_html=True)


def kpi(label, value, sub="", accent=NAVY, bar=None):
    ok = bar is not None and not pd.isna(bar)
    bar_html = (f'<div class="kpi-bar"><div class="kpi-bar-fill" style="width:{min(max(bar, 0), 100):.1f}%;'
                f'background:{accent};"></div></div>') if ok else ''
    return (f'<div class="kpi" title="{esc(KPI_TIPS.get(label, ""))}" style="--accent:{accent}">'
            f'<div class="kpi-label">{esc(label)}</div>'
            f'<div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div>{bar_html}</div>')


def kpi_row(cards):
    st.markdown(f'<div class="kpi-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def tone(value, good, ok):
    if pd.isna(value):
        return SLATE
    return GREEN if value >= good else (AMBER if value >= ok else RED)


def empty_state(title, body):
    st.markdown(f'<div class="empty-state"><b>{esc(title)}</b>{esc(body)}</div>', unsafe_allow_html=True)


def short(s, n=22):
    s = str(s)
    return s if len(s) <= n else s[:n - 1].rstrip() + "…"


def fmt_pct(v):
    return "—" if pd.isna(v) else f"{v:.1f}%"


def fnum(v, fmt=".1f"):
    return "—" if pd.isna(v) else format(v, fmt)


def frank(v):
    return "—" if pd.isna(v) else f"#{int(v)}"


def whatsapp_url(phone, msg):
    return f"https://wa.me/91{phone}?text={urllib.parse.quote(msg)}"


def search_students(df, q, limit=None):
    """Match by ID (substring), name tokens (all must match), or contact; falls back to fuzzy name match."""
    q = q.strip().lower()
    if not q:
        return df
    tokens = q.split()
    name_l = df['Name'].str.lower()
    mask = df['ID'].str.lower().str.contains(re.escape(q), na=False) | \
        df['Contact'].str.contains(re.escape(q), na=False)
    tok_mask = pd.Series(True, index=df.index)
    for t in tokens:
        tok_mask &= name_l.str.contains(re.escape(t), na=False)
    res = df[mask | tok_mask]
    if res.empty and not q.isdigit():
        names = name_l.tolist()
        close = set(difflib.get_close_matches(q, names, n=10, cutoff=0.6))
        # also fuzzy-match individual name words (handles single-word typos)
        for i, nm in enumerate(names):
            for w in nm.split():
                if all(difflib.SequenceMatcher(None, t, w).ratio() > 0.8 for t in tokens[:1]):
                    close.add(nm)
        res = df[name_l.isin(close)]
    return res.head(limit) if limit else res


def open_student(sid):
    st.session_state.sel_id = sid
    st.session_state.view = "Student Dossier"
    st.session_state._last_view = "Student Dossier"


def parent_note(r):
    return (f"Dear Parent, this is an update from the Academic Mentorship Cell regarding {r.Name} "
            f"({r.ID}, {r.Section}). Current attendance is {fmt_pct(r.Att_Pct)} against the required "
            f"{att_thr:.0f}%, and aggregate marks are {fmt_pct(r.Marks_Pct)}. "
            f"Please contact the class mentor at the earliest to discuss an improvement plan.")


def render_cards(df, prefix, step=12):
    """Touch-friendly student cards (shown instead of wide tables on phones)."""
    df = df.reset_index(drop=True)
    n_key = f"cards_n_{prefix}"
    n = st.session_state.get(n_key, step)
    for i, r in enumerate(df.head(n).itertuples(index=False)):
        extra = ""
        if prefix == "int":
            bits = []
            if r.Lec_Needed:
                bits.append(f"<span class='badge badge-info'>Needs {int(r.Lec_Needed)} lectures to recover</span>")
            if r.PTM_Status != "Not met":
                bits.append(f"<span class='badge badge-muted'>PTM: {esc(r.PTM_Status)}</span>")
            extra = f"<div class='scard-extra'>{' '.join(bits)}</div>" if bits else ""
        with st.container(border=True, key=f"cb_{prefix}_{i}"):
            st.markdown(
                f"""<div class="scard"><div class="scard-top"><div><div class="scard-name">{esc(r.Name)}</div>
                <div class="scard-sub">{esc(r.ID)} · {esc(r.Section)}</div></div>
                <span class="badge {r.Badge_Cls}">{esc(r.Risk)}</span></div>
                <div class="scard-stats">
                  <div><span>Attendance</span><b style="color:{tone(r.Att_Pct, att_thr, att_thr - 15)}">{fmt_pct(r.Att_Pct)}</b></div>
                  <div><span>Aggregate</span><b style="color:{tone(r.Marks_Pct, 60, pass_thr)}">{fmt_pct(r.Marks_Pct)}</b></div>
                  <div><span>Below pass</span><b>{int(r.Failed_Subj)} subj</b></div>
                </div>{extra}</div>""", unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            c1.button("Open dossier", key=f"open_{prefix}_{r.ID}", icon=":material/person:", type="primary",
                      width="stretch", on_click=open_student, args=(r.ID,))
            if r.Contact != "N/A":
                url = getattr(r, "WhatsApp", None) or whatsapp_url(r.Contact, parent_note(r))
                c2.link_button("WhatsApp", url, icon=":material/chat:", width="stretch")
    left = len(df) - n
    if left > 0:
        st.button(f"Show {min(step, left)} more ({left} remaining)", key=f"more_{prefix}", width="stretch",
                  on_click=lambda: st.session_state.__setitem__(n_key, n + step))


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — DATA SOURCES & POLICY
# ══════════════════════════════════════════════════════════════════════════════
header_slot = st.container()      # filled in once the data is loaded
policy_slot = st.container()
with policy_slot:
    with st.expander("Policy rules", icon=":material/tune:"):
        _p1, _p2 = st.columns(2)
        att_thr = _p1.slider("Attendance requirement (%)", 50, 90, 75, 5, key="att_thr_s")
        pass_thr = _p2.slider("Pass mark (%)", 30, 60, 40, 5, key="pass_thr_s")
        st.caption("Set the attendance requirement and pass mark. Changes instantly recalculate risk bands, filters, charts and the action list.")

with st.sidebar:
    st.markdown("### Data sources")
    u_marks = st.file_uploader("Marks workbook (.xlsx)", type=["xlsx"], key="u_marks")
    u_att = st.file_uploader("Attendance workbook (.xlsx)", type=["xlsx"], key="u_att")
    term_year = st.number_input("Academic year of attendance dates", 2020, 2035, 2026, 1, key="term_year")

    st.caption("Attendance requirement and pass mark are on the main page under **Policy rules**.")

marks_bytes = u_marks.getvalue() if u_marks else (open(DEFAULT_MARKS, 'rb').read() if os.path.exists(DEFAULT_MARKS) else None)
att_bytes = u_att.getvalue() if u_att else (open(DEFAULT_ATT, 'rb').read() if os.path.exists(DEFAULT_ATT) else None)

if marks_bytes is None:
    st.markdown("## PTM Analytics Portal")
    empty_state("No marks workbook found",
                " Open the sidebar (› at top-left) and upload the marks workbook to begin.")
    st.stop()

try:
    base, df_m, df_detail, meta = load_raw(marks_bytes, att_bytes, int(term_year))
except Exception as e:
    st.error(f"Could not read the uploaded workbook(s): {e}")
    st.stop()

if base.empty:
    empty_state("No student rows detected", " Check that the marks sheet has ID and Name columns.")
    st.stop()

att_thr, pass_thr = float(att_thr), float(pass_thr)
mdf = classify(base, att_thr, pass_thr, meta['subj_max'])
records = load_records()
mdf['PTM_Status'] = mdf['ID'].map(lambda i: records.get(i, {}).get('status', 'Not met'))
id_index = {sid: i for i, sid in enumerate(mdf['ID'])}

with st.sidebar:
    met = (mdf['PTM_Status'].isin(['Met', 'Follow-up required'])).sum()
    st.markdown("### PTM progress")
    st.progress(met / len(mdf), text=f"{met} of {len(mdf)} families consulted")
    fu = (mdf['PTM_Status'] == 'Follow-up required').sum()
    if fu:
        st.caption(f"{fu} follow-up(s) pending")
    if records:
        rec_df = pd.DataFrame([{'ID': k, **v} for k, v in records.items()])
        rec_df = rec_df.merge(mdf[['ID', 'Name', 'Section', 'Contact']], on='ID', how='left')
        st.download_button("Export PTM log (CSV)", rec_df.to_csv(index=False).encode('utf-8'),
                           file_name=f"PTM_Log_{datetime.now():%Y%m%d}.csv", mime="text/csv",
                           icon=":material/download:", width="stretch", key="exp_log")
    if PERSIST:
        st.caption(f"Records saved to `{os.path.basename(RECORDS_PATH)}`")
    else:
        st.caption("Records are kept in this browser session only. Export the PTM log before closing the tab, "
                   "and restore it next time below.")
        u_log = st.file_uploader("Restore PTM log (CSV)", type=["csv"], key="u_log")
        if u_log is not None and st.session_state.get("_log_loaded") != u_log.file_id:
            try:
                log = pd.read_csv(io.BytesIO(u_log.getvalue()), dtype=str).fillna("")
                for _, lr in log.iterrows():
                    if lr.get("ID"):
                        records[str(lr["ID"])] = {k: (lr.get(k) or None) for k in
                                                  ("status", "attendee", "date", "follow_up", "remarks", "updated")}
                        records[str(lr["ID"])]["status"] = records[str(lr["ID"])]["status"] or "Not met"
                        records[str(lr["ID"])]["remarks"] = records[str(lr["ID"])]["remarks"] or ""
                        records[str(lr["ID"])]["attendee"] = records[str(lr["ID"])]["attendee"] or "—"
                st.session_state["_log_loaded"] = u_log.file_id
                st.rerun()
            except Exception as e:
                st.error(f"Could not read that log: {e}")

    n_nomarks = int(mdf['Grand_Total'].isna().sum())
    n_noatt = int((~mdf['In_Att_Sheet']).sum())
    n_drop = int(mdf['Dropped_Out'].sum())
    n_disc = int(((mdf['Att_Official'] - mdf['Att_Calc']).abs() > 5).sum())
    n_phone = int((mdf['Contact'] != 'N/A').sum())
    with st.expander("Data health", icon=":material/verified:"):
        st.markdown(
            f"- **{len(mdf)}** students read from the marks sheet\n"
            f"- **{len(mdf) - n_noatt}** matched in the attendance workbook"
            f"{f' ({n_noatt} not found)' if n_noatt else ''}\n"
            f"- **{n_phone}** have a parent phone number\n"
            f"- Attendance counted from the attendance workbook up to **{meta['att_cutoff'] or 'all dates'}**; cancelled / future classes ignored\n"
            f"- **{n_nomarks}** have no marks on record (shown as unavailable, not zero)\n"
            f"- **{n_drop}** marked drop-out\n"
            f"- **{n_disc}** where the marks-sheet attendance % differs from the workbook count by more than 5 points (workbook is used)\n"
            f"- **{meta['att_orphans']}** attendance rows have no marks record (ignored)")


# ══════════════════════════════════════════════════════════════════════════════
# HEADER, GLOBAL SEARCH, NAVIGATION, FILTERS
# ══════════════════════════════════════════════════════════════════════════════
header_slot.markdown(f"""
<div class="portal-header">
  <div class="brand">
    {f'<img class="brand-logo" src="data:image/png;base64,{LOGO_B64}" alt="Techno India University">' if LOGO_B64 else '<div class="brand-mark">PTM</div>'}
    <div class="brand-divider"></div>
    <div>
      <div class="portal-title">Academic Performance &amp; PTM Portal</div>
      <div class="portal-subtitle">Parent–teacher meeting preparation, student advising and intervention tracking</div>
    </div>
  </div>
  <div class="meta-row">
    <span class="meta-chip" title="Students read from the marks sheet.">Students <b>{len(mdf)}</b></span>
    <span class="meta-chip" title="Class sections found in the attendance workbook.">Sections <b>{mdf['Section'].nunique()}</b></span>
    <span class="meta-chip" title="Maximum total marks, read from the marks sheet header.">Max score <b>{int(meta['max_marks'])}</b></span>
    <span class="meta-chip" title="Policy attendance requirement (change in the sidebar).">Attendance rule <b>{att_thr:.0f}%</b></span>
    <span class="meta-chip" title="Policy pass mark (change in the sidebar).">Pass mark <b>{pass_thr:.0f}%</b></span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Session defaults ─────────────────────────────────────────────────────────
st.session_state.setdefault('view', 'Overview')
st.session_state.setdefault('_last_view', 'Overview')
st.session_state.setdefault('sel_id', None)
FILTER_DEFAULTS = {'f_stream': 'All streams', 'f_section': 'All sections', 'f_att': 'Any attendance',
                   'f_perf': 'Any standing', 'f_risk': 'Any status', 'f_ptm': 'Any PTM status'}
for k, v in FILTER_DEFAULTS.items():
    st.session_state.setdefault(k, v)


def _nav_guard():
    if st.session_state.view is None:
        st.session_state.view = st.session_state._last_view
    st.session_state._last_view = st.session_state.view


def _jump_to_student():
    sid = st.session_state.get('finder')
    if sid:
        open_student(sid)
    st.session_state.finder = None


def _reset_filters():
    for k, v in FILTER_DEFAULTS.items():
        st.session_state[k] = v


labels = {r.ID: f"{r.Name}  ·  {r.ID}  ·  {r.Section}" for r in mdf[['ID', 'Name', 'Section']].itertuples()}

nav_col, find_col = st.columns([3, 2], vertical_alignment="bottom")
with nav_col:
    st.segmented_control("View", VIEWS, key="view", on_change=_nav_guard,
                         format_func=lambda v: f"{VIEW_ICONS[v]} {VIEW_LABELS[v]}", label_visibility="collapsed")
with find_col:
    st.selectbox("Find a student", options=list(labels.keys()), index=None, key="finder",
                 format_func=lambda i: labels.get(i, i), on_change=_jump_to_student,
                 placeholder="Search any student by name or roll number…",
                 label_visibility="collapsed")

view = st.session_state.view or 'Overview'

# ── Filters ──────────────────────────────────────────────────────────────────
try:
    _ua = st.context.headers.get("User-Agent", "") or ""
except Exception:
    _ua = ""
IS_MOBILE_UA = bool(re.search(r"Mobi|Android|iPhone|iPad", _ua))
with st.expander("Filters", icon=":material/filter_list:", expanded=not IS_MOBILE_UA):
    f1, f2, f3, f4, f5, f6, f7 = st.columns([1.2, 1.2, 1.3, 1.3, 1.2, 1.3, 0.7], vertical_alignment="bottom")
    with f1:
        streams = ["All streams"] + sorted(mdf['Stream'].dropna().unique().tolist())
        if st.session_state.f_stream not in streams:
            st.session_state.f_stream = "All streams"
        sel_stream = st.selectbox("Stream", streams, key="f_stream")
    scope = mdf[mdf['Stream'] == sel_stream] if sel_stream != "All streams" else mdf
    with f2:
        secs = ["All sections"] + sorted(scope['Section'].dropna().unique().tolist())
        if st.session_state.f_section not in secs:
            st.session_state.f_section = "All sections"
        sel_sec = st.selectbox("Section", secs, key="f_section")
    with f3:
        att_opts = ["Any attendance", f"Critical (<{att_thr - 15:.0f}%)",
                    f"Borderline ({att_thr - 15:.0f}–{att_thr:.0f}%)", f"Eligible (≥{att_thr:.0f}%)"]
        if st.session_state.f_att not in att_opts:
            st.session_state.f_att = "Any attendance"
        sel_att = st.selectbox("Attendance", att_opts, key="f_att")
    with f4:
        perf_opts = ["Any standing", "High achievers (≥75%)", f"Average ({pass_thr:.0f}–75%)",
                     f"Below pass (<{pass_thr:.0f}%)"]
        if st.session_state.f_perf not in perf_opts:
            st.session_state.f_perf = "Any standing"
        sel_perf = st.selectbox("Academic standing", perf_opts, key="f_perf")
    with f5:
        sel_risk = st.selectbox("Risk status", ["Any status"] + RISK_ORDER, key="f_risk")
    with f6:
        sel_ptm = st.selectbox("PTM status", ["Any PTM status"] + PTM_STATUSES, key="f_ptm")
    with f7:
        st.button("Reset", on_click=_reset_filters, icon=":material/restart_alt:", width="stretch",
                  key="btn_reset")

fdf = scope.copy()
if sel_sec != "All sections":
    fdf = fdf[fdf['Section'] == sel_sec]
if sel_att.startswith("Critical"):
    fdf = fdf[fdf['Att_Pct'] < att_thr - 15]
elif sel_att.startswith("Borderline"):
    fdf = fdf[(fdf['Att_Pct'] >= att_thr - 15) & (fdf['Att_Pct'] < att_thr)]
elif sel_att.startswith("Eligible"):
    fdf = fdf[fdf['Att_Pct'] >= att_thr]
if sel_perf.startswith("High"):
    fdf = fdf[fdf['Marks_Pct'] >= 75]
elif sel_perf.startswith("Average"):
    fdf = fdf[(fdf['Marks_Pct'] >= pass_thr) & (fdf['Marks_Pct'] < 75)]
elif sel_perf.startswith("Below"):
    fdf = fdf[fdf['Marks_Pct'] < pass_thr]
if sel_risk != "Any status":
    fdf = fdf[fdf['Risk'] == sel_risk]
if sel_ptm != "Any PTM status":
    fdf = fdf[fdf['PTM_Status'] == sel_ptm]

active_filters = sum(st.session_state[k] != v for k, v in FILTER_DEFAULTS.items())
scope_title = (sel_sec if sel_sec != "All sections" else
               sel_stream if sel_stream != "All streams" else "All students")
if active_filters:
    st.caption(f"{len(fdf)} of {len(mdf)} students match {active_filters} active filter(s).")


# ══════════════════════════════════════════════════════════════════════════════
# VIEW: OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
def render_overview(sub):
    if sub.empty:
        empty_state("No students in scope", " Broaden or reset the filters above.")
        return
    n = len(sub)
    att_known = sub['Att_Pct'].dropna()
    n_att_def = int((sub['Att_Pct'] < att_thr).sum())
    n_fail = int((sub['Marks_Pct'] < pass_thr).sum())
    n_crit = int((sub['Risk'] == 'Critical').sum())
    met = int(sub['PTM_Status'].isin(['Met', 'Follow-up required']).sum())

    st.markdown(f"## {esc(scope_title)}")
    kpi_row([
        kpi("Students in scope", f"{n}", f"across {sub['Section'].nunique()} section(s)"),
        kpi("Mean attendance", fmt_pct(att_known.mean()) if len(att_known) else "—",
            f"median {fmt_pct(att_known.median())}" if len(att_known) else "no attendance data",
            tone(att_known.mean(), att_thr, att_thr - 15), att_known.mean() if len(att_known) else None),
        kpi("Mean aggregate", f"{sub['Marks_Pct'].mean():.1f}%",
            f"{sub['Grand_Total'].mean():.0f} / {int(meta['max_marks'])} average score",
            tone(sub['Marks_Pct'].mean(), 60, pass_thr), sub['Marks_Pct'].mean()),
        kpi("Below attendance rule", f"{n_att_def}", f"{n_att_def / n * 100:.0f}% of students under {att_thr:.0f}%",
            RED if n_att_def else GREEN, n_att_def / n * 100),
        kpi("Below pass mark", f"{n_fail}", f"{n_fail / n * 100:.0f}% under {pass_thr:.0f}% aggregate",
            RED if n_fail else GREEN, n_fail / n * 100),
        kpi("PTM consultations", f"{met}/{n}", f"{n_crit} critical case(s) in scope", BLUE, met / n * 100),
    ])

    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        section_heading("Risk profile by section", "Number of students in each risk band")
        rp = sub.groupby(['Section', 'Risk']).size().reset_index(name='Students')
        order = sub.groupby('Section')['Risk'].apply(lambda s: (s == 'Critical').mean()).sort_values().index.tolist()
        fig = px.bar(rp, y='Section', x='Students', color='Risk', orientation='h',
                     color_discrete_map=RISK_COLORS, category_orders={'Risk': RISK_ORDER, 'Section': order})
        fig.update_traces(marker_line_width=0)
        fig.update_yaxes(title=None, type='category')
        show_fig(style_fig(fig, max(260, 30 * len(order) + 60), legend=True))
    with c2:
        section_heading("Attendance distribution", f"Dashed line marks the {att_thr:.0f}% requirement")
        fig = px.histogram(sub.dropna(subset=['Att_Pct']), x='Att_Pct', nbins=20,
                           color_discrete_sequence=[NAVY], labels={'Att_Pct': 'Attendance (%)'})
        fig.add_vline(x=att_thr, line_dash="dash", line_color=RED, line_width=1.5)
        fig.update_traces(marker_line_width=0)
        fig.update_yaxes(title="Students")
        show_fig(style_fig(fig, max(260, 30 * sub['Section'].nunique() + 60)))

    section_heading("Section benchmark", "Select a row to filter the dashboard to that section")
    bench = (sub.groupby('Section')
             .agg(Stream=('Stream', 'first'), Students=('ID', 'count'),
                  Attendance=('Att_Pct', 'mean'), Aggregate=('Marks_Pct', 'mean'),
                  Below_Att=('Att_Pct', lambda s: (s < att_thr).mean() * 100),
                  Below_Pass=('Marks_Pct', lambda s: (s < pass_thr).mean() * 100),
                  Critical=('Risk', lambda s: (s == 'Critical').sum()),
                  Consulted=('PTM_Status', lambda s: s.isin(['Met', 'Follow-up required']).sum()))
             .reset_index().sort_values('Aggregate', ascending=False))
    st.session_state._bench_secs = bench['Section'].tolist()

    def _bench_pick():
        rows = st.session_state.bench_tbl.selection.rows
        if rows:
            sec = st.session_state._bench_secs[rows[0]]
            st.session_state.f_stream = "All streams"
            st.session_state.f_section = sec

    st.dataframe(
        bench, hide_index=True, width="stretch", key="bench_tbl", on_select=_bench_pick,
        selection_mode="single-row",
        column_config={
            "Attendance": st.column_config.ProgressColumn("Avg attendance", format="%.1f%%", min_value=0, max_value=100, help="Mean attendance in the section (attendance workbook)."),
            "Aggregate": st.column_config.ProgressColumn("Avg aggregate", format="%.1f%%", min_value=0, max_value=100, help="Mean total marks % of students with marks."),
            "Below_Att": st.column_config.NumberColumn(f"< {att_thr:.0f}% attendance", format="%.0f%%"),
            "Below_Pass": st.column_config.NumberColumn(f"< {pass_thr:.0f}% marks", format="%.0f%%"),
            "Critical": st.column_config.NumberColumn("Critical", help="Students in the highest-risk band."),
            "Consulted": st.column_config.NumberColumn("PTM done", help="Parent meetings completed or needing follow-up."),
        })

    section_heading("Attendance vs. academic performance",
                    "Each dot is a student — hover for details, click to open the dossier")
    pts = sub.dropna(subset=['Att_Pct', 'Marks_Pct'])
    fig = px.scatter(pts, x='Att_Pct', y='Marks_Pct', color='Risk', color_discrete_map=RISK_COLORS,
                     category_orders={'Risk': RISK_ORDER}, custom_data=['ID', 'Name', 'Section'],
                     labels={'Att_Pct': 'Attendance (%)', 'Marks_Pct': 'Aggregate (%)'})
    fig.update_traces(marker=dict(size=7, opacity=0.75, line=dict(width=0.5, color="#ffffff")),
                      hovertemplate="<b>%{customdata[1]}</b><br>%{customdata[0]} · %{customdata[2]}"
                                    "<br>Attendance %{x:.1f}% · Aggregate %{y:.1f}%<extra></extra>")
    if len(pts) > 2 and pts['Att_Pct'].nunique() > 1:
        cf = np.polyfit(pts['Att_Pct'], pts['Marks_Pct'], 1)
        xs = np.linspace(pts['Att_Pct'].min(), pts['Att_Pct'].max(), 50)
        r = np.corrcoef(pts['Att_Pct'], pts['Marks_Pct'])[0, 1]
        fig.add_trace(go.Scatter(x=xs, y=cf[0] * xs + cf[1], mode='lines', name=f'Trend (r = {r:.2f})',
                                 line=dict(color=SLATE, dash='dot', width=1.5), hoverinfo='skip'))
    fig.add_vline(x=att_thr, line_dash="dash", line_color=RED, line_width=1)
    fig.add_hline(y=pass_thr, line_dash="dash", line_color=RED, line_width=1)
    ev = st.plotly_chart(style_fig(fig, 420, legend=True, static=False), width="stretch", config=CHART_CFG, theme=None,
                         on_select="rerun", selection_mode="points", key="ov_scatter")
    try:
        pt = ev.selection.points
        if pt and pt[0].get('customdata'):
            open_student(pt[0]['customdata'][0])
            st.rerun()
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
# VIEW: STUDENT DOSSIER
# ══════════════════════════════════════════════════════════════════════════════
def build_message(s, best, worst, remarks):
    status = ("Compliant" if s['Att_Pct'] >= att_thr else
              "Below requirement" if s['Att_Pct'] >= att_thr - 15 else "Critically low") \
        if pd.notna(s['Att_Pct']) else "Not available"
    lines = [
        f"*Academic Progress Update — {datetime.now():%d %b %Y}*",
        "",
        f"Student: {s['Name']} (ID: {s['ID']})",
        f"Section: {s['Section']}",
        "",
        "*Summary*",
        f"• Attendance: {fmt_pct(s['Att_Pct'])} ({status}; requirement {att_thr:.0f}%)",
        f"• Aggregate: {fnum(s['Grand_Total'])}/{int(s['Max_Marks'])} ({fmt_pct(s['Marks_Pct'])})",
        f"• Section rank: {frank(s['Section_Rank'])} of {int(s['Section_Size'])}",
    ]
    if s['Lec_Needed']:
        lines.append(f"• Needs to attend the next {s['Lec_Needed']} lectures without absence to reach {att_thr:.0f}%")
    if best:
        lines += ["", "*Observations*", f"• Strongest subject: {best[0]} ({best[1]:.1f})",
                  f"• Needs attention: {worst[0]} ({worst[1]:.1f})"]
    lines += ["", "*Faculty remarks*",
              remarks.strip() or "Please ensure regular attendance and monitor daily study at home.",
              "", "— Academic Mentorship Cell"]
    return "\n".join(lines)


def build_report_html(s, sub_rows, sub_att, rec):
    srows = "".join(
        f"<tr><td>{esc(n)}</td><td class='r'>{v:.1f}</td><td class='r'>{a:.1f}</td>"
        f"<td class='r' style='color:{'#059669' if v >= a else '#dc2626'}'>{v - a:+.1f}</td></tr>"
        for n, v, a in sub_rows)
    arows = "".join(f"<tr><td>{esc(n)}</td><td class='r'>{a}/{t}</td><td class='r'>{a / t * 100:.0f}%</td></tr>"
                    for n, (a, t) in sub_att)
    remarks = esc(rec.get('remarks', '') or '—').replace("\n", "<br>")
    return f"""<!doctype html><html><head><meta charset="utf-8"><title>PTM Report — {esc(s['Name'])}</title>
<style>
body{{font-family:Inter,Arial,sans-serif;color:#0f172a;max-width:820px;margin:32px auto;padding:0 24px;font-size:13px}}
h1{{font-size:20px;margin:0}} h2{{font-size:13px;text-transform:uppercase;letter-spacing:.05em;color:#1e3a5f;
border-bottom:1px solid #e2e8f0;padding-bottom:4px;margin-top:22px}}
.head{{display:flex;justify-content:space-between;border-bottom:2px solid #1e3a5f;padding-bottom:10px}}
.muted{{color:#64748b}} table{{width:100%;border-collapse:collapse}} td,th{{padding:6px 8px;border-bottom:1px solid #eef2f7;text-align:left}}
th{{background:#f8fafc;font-size:11px;text-transform:uppercase;color:#475569}} .r{{text-align:right}}
.grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}} .box{{border:1px solid #e2e8f0;border-radius:6px;padding:10px}}
.box b{{display:block;font-size:18px;margin-top:4px}} .sig{{display:flex;justify-content:space-between;margin-top:60px}}
.sig div{{border-top:1px solid #94a3b8;width:40%;padding-top:4px;text-align:center;color:#475569}}
@media print{{body{{margin:0}}}}
</style></head><body>
<div class="head"><div>{f'<img src="data:image/png;base64,{LOGO_B64}" style="height:44px;margin-bottom:8px">' if LOGO_B64 else ''}<h1>Parent–Teacher Meeting Report</h1><div class="muted">Academic Mentorship Cell</div></div>
<div class="muted" style="text-align:right">Generated {datetime.now():%d %b %Y}<br>PTM status: {esc(rec.get('status', 'Not met'))}</div></div>
<h2>Student</h2>
<table><tr><td><b>Name</b></td><td>{esc(s['Name'])}</td><td><b>Roll ID</b></td><td>{esc(s['ID'])}</td></tr>
<tr><td><b>Section</b></td><td>{esc(s['Section'])}</td><td><b>Stream</b></td><td>{esc(s['Stream'])}</td></tr>
<tr><td><b>Parent contact</b></td><td>{esc(s['Contact'])}</td><td><b>Risk status</b></td><td>{esc(s['Risk'])}</td></tr></table>
<h2>Summary</h2><div class="grid">
<div class="box">Attendance<b>{fmt_pct(s['Att_Pct'])}</b><span class="muted">requirement {att_thr:.0f}%</span></div>
<div class="box">Aggregate<b>{fnum(s['Grand_Total'])}</b><span class="muted">of {int(s['Max_Marks'])} ({fmt_pct(s['Marks_Pct'])})</span></div>
<div class="box">Section rank<b>{frank(s['Section_Rank'])}</b><span class="muted">of {int(s['Section_Size'])}</span></div>
<div class="box">Cohort rank<b>{frank(s['Batch_Rank'])}</b><span class="muted">of {int(s['Batch_Size'])}</span></div></div>
<h2>Subject performance</h2><table><tr><th>Subject</th><th class="r">Score</th><th class="r">Section avg</th><th class="r">Difference</th></tr>{srows}</table>
{f'<h2>Subject attendance</h2><table><tr><th>Subject</th><th class="r">Attended</th><th class="r">Rate</th></tr>{arows}</table>' if arows else ''}
<h2>Faculty remarks &amp; agreed actions</h2><p>{remarks}</p>
{f"<p class='muted'>Follow-up date: {esc(rec['follow_up'])}</p>" if rec.get('follow_up') else ''}
<div class="sig"><div>Faculty Mentor</div><div>Parent / Guardian</div></div>
</body></html>"""


def render_dossier(sub):
    ids = sub['ID'].tolist()
    sid = st.session_state.sel_id
    if sid not in id_index:
        sid = ids[0] if ids else None
        st.session_state.sel_id = sid
    if sid is None:
        empty_state("No student selected", " Use the search box at the top right, or broaden the filters.")
        return
    s = mdf.iloc[id_index[sid]]
    rec = records.get(sid, {})

    # ── Navigation strip ─────────────────────────────────────────────────────
    pos = ids.index(sid) if sid in ids else None
    with st.container(key="dnav"):
        n1, n2, n3, n4 = st.columns([1, 1, 3, 1.4], vertical_alignment="center")
    with n1:
        st.button("Previous", icon=":material/chevron_left:", width="stretch", key="btn_prev",
                  disabled=not ids or pos == 0,
                  on_click=lambda: st.session_state.update(
                      sel_id=ids[(pos - 1) if pos is not None else 0]))
    with n2:
        st.button("Next", icon=":material/chevron_right:", icon_position="right", width="stretch", key="btn_next",
                  disabled=not ids or (pos is not None and pos >= len(ids) - 1),
                  on_click=lambda: st.session_state.update(
                      sel_id=ids[(pos + 1) if pos is not None else 0]))
    with n3:
        st.caption(f"Student {pos + 1} of {len(ids)} in current filter" if pos is not None
                   else "This student is outside the current filter — Previous/Next will jump into the filtered list.")

    # ── Subject data ─────────────────────────────────────────────────────────
    sub_scores = dict(s['Subj'])
    sec_ids = mdf.loc[mdf['Section'] == s['Section'], 'ID']
    sec_m = df_m[df_m['ID_clean'].isin(sec_ids)]
    sec_avg = {sc: float(sec_m[sc].mean()) for sc in sub_scores}
    ranked = sorted(sub_scores.items(), key=lambda kv: kv[1], reverse=True)
    best = ranked[0] if ranked else None
    worst = ranked[-1] if ranked else None
    cut = pass_thr / 100 * meta['subj_max']
    failing = [k for k, v in ranked if v < cut]
    sub_att = sorted([(k, v) for k, v in s['Sub_A'].items() if v[1] > 0],
                     key=lambda kv: kv[1][0] / kv[1][1])

    with n4:
        report = build_report_html(s, [(k, v, sec_avg[k]) for k, v in ranked], sub_att, rec)
        st.download_button("Printable report", report.encode('utf-8'),
                           file_name=f"PTM_Report_{sid}.html", mime="text/html",
                           icon=":material/print:", width="stretch", key="btn_report")

    # ── Profile card ─────────────────────────────────────────────────────────
    initials = "".join(w[0].upper() for w in str(s['Name']).split()[:2]) or "?"
    ptm_cls = {'Met': 'badge-safe', 'Follow-up required': 'badge-warn',
               'Parent absent': 'badge-danger'}.get(s['PTM_Status'], 'badge-muted')
    st.markdown(f"""
    <div class="profile-card">
      <div class="profile-initials">{esc(initials)}</div>
      <div>
        <div class="profile-name">{esc(s['Name'])}</div>
        <div class="profile-meta">
          <span>Roll ID <b>{esc(s['ID'])}</b></span>
          <span>Section <b>{esc(s['Section'])}</b></span>
          <span>Stream <b>{esc(s['Stream'])}</b></span>
          <span>Parent contact <b>{esc(s['Contact'])}</b></span>
        </div>
      </div>
      <div class="profile-badges">
        <span class="badge {s['Badge_Cls']}" title="Risk band from attendance (requirement {att_thr:.0f}%) and marks (pass {pass_thr:.0f}%).">{esc(s['Risk'])}</span>
        <span class="badge {ptm_cls}" title="Status of the parent meeting for this student.">PTM: {esc(s['PTM_Status'])}</span>
        {'<span class="badge badge-danger" title="Marked DROP OUT in the attendance sheet.">Drop-out</span>' if s['Dropped_Out'] else ''}
      </div>
    </div>""", unsafe_allow_html=True)

    notes = []
    if s['Dropped_Out']:
        notes.append("This student is marked DROP OUT in the attendance workbook.")
    if pd.isna(s['Marks_Pct']):
        notes.append("No marks are recorded for this student, so marks-based figures and ranks are hidden.")
    if not s['In_Att_Sheet']:
        notes.append("Not found in the attendance workbook: parent contact and lecture-level details are unavailable. "
                     "The attendance % shown is the fallback value from the marks sheet.")
    elif pd.notna(s['Att_Official']) and pd.notna(s['Att_Pct']) and abs(s['Att_Pct'] - s['Att_Official']) > 5:
        notes.append(f"Attendance is counted from the attendance workbook ({s['Att_Pct']:.1f}% = "
                     f"{int(s['Attended'])}/{int(s['Total_Lec'])} lectures). The marks sheet lists {s['Att_Official']:.1f}% "
                     "for the same student; that figure is not used.")
    for n_ in notes:
        st.info(n_, icon=":material/info:")

    if s['Contact'] not in ('N/A', ''):
        with st.container(key="qa_row"):
            q1, q2 = st.columns(2)
            q1.link_button("WhatsApp parent", whatsapp_url(s['Contact'], build_message(s, best, worst, rec.get('remarks', ''))),
                           type="primary", icon=":material/chat:", width="stretch")
            q2.link_button("Call parent", f"tel:+91{s['Contact']}", icon=":material/call:", width="stretch")

    # ── KPIs ─────────────────────────────────────────────────────────────────
    att_clr = tone(s['Att_Pct'], att_thr, att_thr - 15)
    att_sub = (f"{int(s['Attended'])} of {int(s['Total_Lec'])} lectures attended"
               if pd.notna(s['Attended']) and pd.notna(s['Total_Lec']) and s['Total_Lec'] else "lecture counts unavailable")
    sec_mean = mdf.loc[mdf['Section'] == s['Section'], 'Marks_Pct'].mean()
    has_marks = pd.notna(s['Marks_Pct'])
    delta = (s['Marks_Pct'] - sec_mean) if has_marks else np.nan
    delta_txt = (f"(<span style='color:{GREEN if delta >= 0 else RED}'>{delta:+.1f} vs section</span>)"
                 if has_marks and pd.notna(delta) else "")
    kpi_row([
        kpi("Attendance", fmt_pct(s['Att_Pct']), att_sub, att_clr,
            s['Att_Pct'] if pd.notna(s['Att_Pct']) else None),
        kpi("Aggregate score", fnum(s['Grand_Total']),
            f"of {int(s['Max_Marks'])} · {fmt_pct(s['Marks_Pct'])} {delta_txt}" if has_marks else "no marks on record",
            tone(s['Marks_Pct'], 60, pass_thr), s['Marks_Pct'] if has_marks else None),
        kpi("Section rank", frank(s['Section_Rank']), f"of {int(s['Section_Size'])} ranked in {esc(s['Section'])}", "#7c3aed"),
        kpi("Cohort percentile",
            f"{s['Percentile']:.0f}<span style='font-size:1rem'>th</span>" if has_marks else "—",
            f"rank {frank(s['Batch_Rank'])} of {int(s['Batch_Size'])}", BLUE, s['Percentile'] if has_marks else None),
        kpi("Subjects below pass", f"{len(failing)}", f"of {len(sub_scores)} subjects under {cut:.0f} marks",
            RED if failing else GREEN),
    ])

    # ── Talking points & recommended actions ─────────────────────────────────
    points, actions = [], []
    if best:
        top2 = ", ".join(f"<b>{esc(k)}</b> ({v:.0f})" for k, v in ranked[:2])
        points.append((GREEN, f"Strongest subjects: {top2}."))
    if failing:
        points.append((RED, f"Below pass mark in <b>{len(failing)}</b> subject(s): "
                            + ", ".join(f"<b>{esc(k)}</b> ({sub_scores[k]:.0f})" for k in failing[:4]) + "."))
        actions.append(f"Arrange remedial support / doubt-clearing sessions for {esc(failing[0])}.")
    elif worst:
        points.append((AMBER, f"Lowest-scoring subject: <b>{esc(worst[0])}</b> ({worst[1]:.0f}, "
                              f"section avg {sec_avg[worst[0]]:.0f})."))
    if pd.notna(s['Att_Pct']):
        if s['Att_Pct'] < att_thr:
            points.append((RED, f"Attendance <b>{s['Att_Pct']:.1f}%</b> is below the {att_thr:.0f}% requirement — "
                                f"risk of exam debarment. Needs <b>{s['Lec_Needed']}</b> consecutive lectures to recover."))
            actions.append("Agree on a daily attendance check-in with the parent; share weekly attendance updates.")
        else:
            points.append((GREEN, f"Attendance <b>{s['Att_Pct']:.1f}%</b> meets the {att_thr:.0f}% requirement."))
    low_att = [(k, a, t) for k, (a, t) in sub_att if t >= 5 and a / t * 100 < att_thr]
    if low_att:
        points.append((AMBER, "Lowest subject attendance: " +
                       ", ".join(f"<b>{esc(k)}</b> ({a / t * 100:.0f}%)" for k, a, t in low_att[:3]) + "."))
    mon = [(m, v) for m, v in s['Mon_S'].items() if v[1] > 0]
    if len(mon) >= 2:
        first, last = mon[0][1][0] / mon[0][1][1] * 100, mon[-1][1][0] / mon[-1][1][1] * 100
        trend = "improving" if last - first > 3 else "declining" if first - last > 3 else "stable"
        points.append((GREEN if trend == "improving" else RED if trend == "declining" else SLATE,
                       f"Monthly attendance trend is <b>{trend}</b> ({mon[0][0]} {first:.0f}% → {mon[-1][0]} {last:.0f}%)."))
        if trend == "declining":
            actions.append("Discuss reasons for the recent drop in attendance (health, commute, motivation).")
    if s['Marks_Pct'] >= 75 and not failing:
        actions.append("Recognise the strong performance; consider advanced electives or competitions.")
    if not actions:
        actions.append("Maintain current routine; review progress at the next PTM.")

    p1, p2 = st.columns([3, 2], gap="medium")
    with p1:
        items = "".join(f'<div class="insight-item"><span class="dot" style="background:{c}"></span><div>{t}</div></div>'
                        for c, t in points)
        st.markdown(f'<div class="insight-panel"><h4>Talking points</h4>{items}</div>', unsafe_allow_html=True)
    with p2:
        items = "".join(f'<div class="insight-item"><span class="dot" style="background:{BLUE}"></span><div>{a}</div></div>'
                        for a in actions)
        st.markdown(f'<div class="insight-panel"><h4>Suggested actions</h4>{items}</div>', unsafe_allow_html=True)

    st.write("")

    # ── Academic performance ─────────────────────────────────────────────────
    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        section_heading("Subject performance vs. section average")
        if sub_scores:
            names = [k for k, _ in ranked][::-1]
            fig = go.Figure([
                go.Bar(name="Section average", y=[short(n, 34) for n in names], x=[sec_avg[n] for n in names],
                       orientation='h', marker_color=LIGHT, hovertemplate="%{x:.1f}<extra>Section avg</extra>"),
                go.Bar(name="Student", y=[short(n, 34) for n in names], x=[sub_scores[n] for n in names],
                       orientation='h', marker_color=[RED if sub_scores[n] < cut else NAVY for n in names],
                       text=[f"{sub_scores[n]:.0f}" for n in names], textposition="outside",
                       hovertemplate="%{x:.1f}<extra>Student</extra>"),
            ])
            fig.add_vline(x=cut, line_dash="dot", line_color=RED, line_width=1,
                          annotation_text="Pass", annotation_font_color=RED, annotation_font_size=10)
            fig.update_layout(barmode="group", bargroupgap=0.05)
            fig.update_xaxes(range=[0, meta['subj_max'] * 1.1], title=None)
            show_fig(style_fig(fig, max(300, 46 * len(names)), legend=True))
        else:
            empty_state("No subject marks", " This student has no subject-level marks.")

    with c2:
        section_heading("Assessment breakdown")
        det = None
        if df_detail is not None:
            det = df_detail[df_detail['ID_clean'] == sid]
        if det is not None and not det.empty:
            drop = [c for c in det.columns if c == 'ID_clean' or 'code' in str(c).lower()
                    or re.search(r'\b(id|reg|roll|name)\b', str(c).lower()) and 'subject' not in str(c).lower()]
            det = det.drop(columns=drop)
            num_cols = [c for c in det.columns if pd.api.types.is_numeric_dtype(det[c])]
            cfg = {c: st.column_config.NumberColumn(format="%.1f") for c in num_cols}
            tot_c = next((c for c in det.columns if 'total' in str(c).lower()), None)
            if tot_c:
                cfg[tot_c] = st.column_config.ProgressColumn(tot_c, format="%.1f", min_value=0,
                                                             max_value=float(max(100, det[tot_c].max() or 0)))
            name_c = next((c for c in det.columns if 'name' in str(c).lower()), None)
            if name_c:
                cfg[name_c] = st.column_config.TextColumn(name_c, pinned=True, width="medium")
            st.dataframe(det, hide_index=True, width="stretch", column_config=cfg,
                         height=min(40 + 35 * len(det), 460))
        elif sub_scores:
            st.dataframe(pd.DataFrame([{'Subject': k, 'Score': v, 'Section avg': sec_avg[k]} for k, v in ranked]),
                         hide_index=True, width="stretch",
                         column_config={'Score': st.column_config.ProgressColumn(format="%.1f", min_value=0,
                                                                                 max_value=meta['subj_max']),
                                        'Section avg': st.column_config.NumberColumn(format="%.1f")})

    # ── Attendance intelligence ──────────────────────────────────────────────
    st.divider()
    section_heading("Attendance intelligence", "Patterns by weekday and month, plus a recovery calculator")
    a1, a2, a3 = st.columns(3, gap="large")
    with a1:
        st.markdown("**By day of week**")
        day_d = s['Day_S']
        held = [d for d in WEEKDAYS if day_d.get(d, [0, 0])[1] > 0]
        if held:
            rates = [day_d[d][0] / day_d[d][1] * 100 for d in held]
            fig = go.Figure(go.Bar(x=[d[:3] for d in held], y=rates,
                                   marker_color=[tone(r, att_thr, att_thr - 15) for r in rates],
                                   text=[f"{r:.0f}%" for r in rates], textposition="outside",
                                   customdata=[f"{day_d[d][0]}/{day_d[d][1]}" for d in held],
                                   hovertemplate="%{x}: %{y:.1f}% (%{customdata})<extra></extra>"))
            fig.update_yaxes(range=[0, 115], showticklabels=False)
            show_fig(style_fig(fig, 220))
            worst_day = held[int(np.argmin(rates))]
            st.caption(f"Most absences on **{worst_day}s** ({min(rates):.0f}% attended).")
        else:
            st.caption("No dated attendance records.")
    with a2:
        st.markdown("**Monthly trend**")
        if mon:
            rates = [v[0] / v[1] * 100 for _, v in mon]
            fig = go.Figure(go.Scatter(x=[m for m, _ in mon], y=rates, mode='lines+markers+text',
                                       text=[f"{r:.0f}%" for r in rates], textposition="top center",
                                       line=dict(color=NAVY, width=2.5),
                                       marker=dict(size=8, color=[tone(r, att_thr, att_thr - 15) for r in rates]),
                                       hovertemplate="%{x}: %{y:.1f}%<extra></extra>"))
            fig.add_hline(y=att_thr, line_dash="dash", line_color=RED, line_width=1)
            fig.update_yaxes(range=[0, 115], showticklabels=False)
            show_fig(style_fig(fig, 220))
        else:
            st.caption("No monthly attendance records.")
    with a3:
        st.markdown("**Recovery calculator**")
        if pd.notna(s['Attended']) and pd.notna(s['Total_Lec']) and s['Total_Lec']:
            goal = st.slider("Target attendance (%)", 50, 95, int(att_thr), 5, key=f"sim_{sid}")
            cur_a, cur_t = int(s['Attended_Eff']), int(s['Total_Lec'])
            needed = lectures_needed(cur_a, cur_t, goal)
            if needed:
                st.markdown(f"""<div class="sim sim-alert">To reach <b>{goal}%</b>, the student must attend the next
                  <span class="sim-num">{needed} lectures</span>without a single absence
                  (currently {cur_a}/{cur_t}).</div>""", unsafe_allow_html=True)
            else:
                buf = int(np.floor((100 * cur_a - goal * cur_t) / goal)) if goal else 0
                st.markdown(f"""<div class="sim sim-ok">Meets the <b>{goal}%</b> target. Can miss up to
                  <span class="sim-num">{buf} lectures</span>before falling below it.</div>""",
                            unsafe_allow_html=True)
        else:
            st.caption("Lecture-level counts are not available for this student.")

    if sub_att:
        with st.expander(f"Subject-wise attendance register ({len(sub_att)} subjects)", icon=":material/fact_check:"):
            reg = pd.DataFrame([{'Subject': k, 'Attended': a, 'Held': t, 'Rate': a / t * 100,
                                 'Status': att_band(a / t * 100, att_thr)[1]} for k, (a, t) in sub_att])
            st.dataframe(reg, hide_index=True, width="stretch",
                         column_config={'Rate': st.column_config.ProgressColumn("Attendance", format="%.0f%%",
                                                                                min_value=0, max_value=100)})

    # ── PTM record & parent communication ────────────────────────────────────
    st.divider()
    section_heading("PTM record & parent communication",
                    "Saved records persist across sessions and feed the PTM status filter")
    r1, r2 = st.columns(2, gap="large")
    with r1:
        with st.form(key=f"ptm_form_{sid}", border=True):
            fa, fb = st.columns(2)
            status = fa.selectbox("Meeting status", PTM_STATUSES, key=f"ptm_status_{sid}",
                                  index=PTM_STATUSES.index(rec.get('status', 'Not met')))
            attendee = fb.selectbox("Attended by", ["—", "Father", "Mother", "Both parents", "Guardian", "Student only"],
                                    key=f"ptm_att_{sid}",
                                    index=["—", "Father", "Mother", "Both parents", "Guardian", "Student only"]
                                    .index(rec.get('attendee', '—')))
            fc, fd = st.columns(2)
            mdate = fc.date_input("Meeting date", key=f"ptm_date_{sid}", value=date.fromisoformat(rec['date']) if rec.get('date') else date.today())
            fup = fd.date_input("Follow-up by (optional)", key=f"ptm_fup_{sid}",
                                value=date.fromisoformat(rec['follow_up']) if rec.get('follow_up') else None)
            remarks = st.text_area("Remarks & agreed actions", value=rec.get('remarks', ''), height=130,
                                   key=f"ptm_remarks_{sid}",
                                   placeholder="e.g. Parent informed of attendance shortfall. Agreed on daily check-ins "
                                               "and weekend revision for Mathematics.")
            if st.form_submit_button("Save PTM record", type="primary", icon=":material/save:", key="btn_save_ptm"):
                ok = save_record(sid, dict(status=status, attendee=attendee, date=mdate.isoformat(),
                                           follow_up=fup.isoformat() if fup else None, remarks=remarks))
                st.toast("PTM record saved." if ok else "Saved for this session only (file not writable).",
                         icon=":material/check_circle:" if ok else ":material/warning:")
                st.rerun()
        if rec.get('updated'):
            st.caption(f"Last updated {rec['updated']}")

    with r2:
        msg = build_message(s, best, worst, rec.get('remarks', ''))
        st.markdown("**Parent message**")
        st.code(msg, language=None, wrap_lines=True, height=300)
        b1, b2 = st.columns(2)
        if s['Contact'] not in ('N/A', ''):
            b1.link_button(f"Send on WhatsApp", whatsapp_url(s['Contact'], msg), type="primary",
                           icon=":material/chat:", width="stretch")
            b2.link_button(f"Call +91 {s['Contact']}", f"tel:+91{s['Contact']}", icon=":material/call:",
                           width="stretch")
        else:
            st.caption("No parent contact on record — use the copy icon on the message to share it manually.")


# ══════════════════════════════════════════════════════════════════════════════
# VIEW: DIRECTORY (search + browse)
# ══════════════════════════════════════════════════════════════════════════════
def render_directory(sub):
    st.markdown("## Student directory")
    q = st.text_input("Search", key="dir_q", placeholder="Search by name, roll number, or parent phone — "
                      "partial matches and minor typos are fine (e.g. “subham”, “5001”, “niten”)",
                      icon=":material/search:", label_visibility="collapsed")
    res = search_students(sub, q) if q.strip() else sub
    if q.strip() and res.empty:
        empty_state("No matches", f" Nothing in the current filter matches “{q}”. Try fewer letters or reset filters.")
        return
    sort_by = st.pills("Sort by", ["Name", "Roll ID", "Attendance ↑", "Aggregate ↓", "Aggregate ↑", "Risk"],
                       default="Name", key="dir_sort")
    key_map = {"Name": ('Name', True), "Roll ID": ('ID', True), "Attendance ↑": ('Att_Pct', True),
               "Aggregate ↓": ('Marks_Pct', False), "Aggregate ↑": ('Marks_Pct', True)}
    if sort_by == "Risk":
        res = res.assign(_r=res['Risk'].map({k: i for i, k in enumerate(RISK_ORDER)})).sort_values(['_r', 'Att_Pct'])
    elif sort_by in key_map:
        res = res.sort_values(key_map[sort_by][0], ascending=key_map[sort_by][1])

    st.caption(f"{len(res)} student(s). Select a row to open the dossier.")
    table = res[['ID', 'Name', 'Section', 'Contact', 'Att_Pct', 'Marks_Pct', 'Failed_Subj',
                 'Section_Rank', 'Risk', 'PTM_Status']].reset_index(drop=True)
    st.session_state._dir_ids = table['ID'].tolist()

    def _pick():
        rows = st.session_state.dir_tbl.selection.rows
        if rows:
            open_student(st.session_state._dir_ids[rows[0]])

    with st.container(key="tbl_dir"):
        st.dataframe(table, hide_index=True, width="stretch", height=560, key="dir_tbl",
                     on_select=_pick, selection_mode="single-row",
                     column_config={
                         'ID': st.column_config.TextColumn("Roll ID", pinned=True),
                         'Name': st.column_config.TextColumn("Name", pinned=True, width="medium"),
                         'Att_Pct': st.column_config.ProgressColumn("Attendance", format="%.1f%%", min_value=0, max_value=100, help="Attendance % counted from the attendance workbook up to the cut-off date."),
                         'Marks_Pct': st.column_config.ProgressColumn("Aggregate", format="%.1f%%", min_value=0, max_value=100, help="Total marks as a percentage of the maximum."),
                         'Failed_Subj': st.column_config.NumberColumn("Subjects < pass", help="Number of subjects scored under the pass mark."),
                         'Section_Rank': st.column_config.NumberColumn("Sec. rank", format="%d", help="Rank by total marks within the section."),
                         'PTM_Status': st.column_config.TextColumn("PTM", help="Parent-meeting status recorded for this student."),
                     })
    with st.container(key="cards_dir"):
        render_cards(res, "dir")
    st.download_button("Download this list (CSV)", table.to_csv(index=False).encode('utf-8'),
                       file_name=f"Students_{datetime.now():%Y%m%d}.csv", mime="text/csv",
                       icon=":material/download:", key="dl_dir")


# ══════════════════════════════════════════════════════════════════════════════
# VIEW: SUBJECTS
# ══════════════════════════════════════════════════════════════════════════════
def render_subjects(sub):
    st.markdown(f"## Subject analytics — {esc(scope_title)}")
    subj_cols = meta['subj_cols']
    if not subj_cols or sub.empty:
        empty_state("No subject data", " No subject columns or no students in scope.")
        return
    cut = pass_thr / 100 * meta['subj_max']
    sm = df_m[df_m['ID_clean'].isin(sub['ID'])].merge(sub[['ID', 'Section']], left_on='ID_clean', right_on='ID')
    stats = []
    for sc in subj_cols:
        v = sm[sc].dropna()
        if v.empty:
            continue
        stats.append(dict(Subject=sc, Students=len(v), Average=v.mean(), Median=v.median(),
                          Highest=v.max(), Pass_Rate=(v >= cut).mean() * 100, Below_Pass=int((v < cut).sum())))
    st_df = pd.DataFrame(stats).sort_values('Pass_Rate')
    hardest = st_df.iloc[0]
    easiest = st_df.iloc[-1]
    kpi_row([
        kpi("Subjects", f"{len(st_df)}", f"scored out of {meta['subj_max']:.0f} each"),
        kpi("Lowest pass rate", f"{hardest['Pass_Rate']:.0f}%", esc(short(hardest['Subject'], 40)), RED, hardest['Pass_Rate']),
        kpi("Highest pass rate", f"{easiest['Pass_Rate']:.0f}%", esc(short(easiest['Subject'], 40)), GREEN, easiest['Pass_Rate']),
        kpi("Failing grades", f"{int(st_df['Below_Pass'].sum())}", f"subject results under {cut:.0f} marks", AMBER),
    ])

    c1, c2 = st.columns([1, 1], gap="large")
    with c1:
        section_heading("Pass rate by subject", f"Share of students scoring ≥ {cut:.0f}")
        fig = go.Figure(go.Bar(y=[short(s, 34) for s in st_df['Subject']], x=st_df['Pass_Rate'], orientation='h',
                               marker_color=[tone(p, 85, 70) for p in st_df['Pass_Rate']],
                               text=[f"{p:.0f}%" for p in st_df['Pass_Rate']], textposition="outside",
                               customdata=st_df['Subject'],
                               hovertemplate="%{customdata}<br>Pass rate %{x:.1f}%<extra></extra>"))
        fig.update_xaxes(range=[0, 110], title=None)
        show_fig(style_fig(fig, max(320, 38 * len(st_df))))
    with c2:
        section_heading("Score spread by subject", "Box shows middle 50% of scores; dots are outliers")
        long = sm.melt(id_vars=['ID_clean'], value_vars=subj_cols, var_name='Subject', value_name='Score').dropna()
        long['Subject'] = long['Subject'].map(lambda s: short(s, 34))
        order = [short(s, 34) for s in st_df.sort_values('Median')['Subject']]
        fig = px.box(long, y='Subject', x='Score', orientation='h', category_orders={'Subject': order[::-1]},
                     color_discrete_sequence=[NAVY])
        fig.add_vline(x=cut, line_dash="dot", line_color=RED, line_width=1)
        fig.update_traces(marker=dict(size=3, opacity=0.5), line_width=1.2)
        fig.update_yaxes(title=None)
        show_fig(style_fig(fig, max(320, 38 * len(st_df))))

    if sm['Section'].nunique() > 1:
        section_heading("Section × subject averages", "Darker cells indicate higher average scores")
        hm = sm.groupby('Section')[subj_cols].mean().round(1).dropna(how='all')
        hm.columns = [short(c, 22) for c in hm.columns]
        fig = px.imshow(hm, text_auto=".0f", aspect="auto",
                        color_continuous_scale=[[0, "#fef2f2"], [0.5, "#dbeafe"], [1, NAVY]])
        fig.update_layout(coloraxis_showscale=False)
        fig.update_xaxes(side="top", tickangle=-30, title=None)
        fig.update_yaxes(title=None)
        show_fig(style_fig(fig, max(300, 34 * len(hm) + 120)))

    section_heading("Subject summary")
    st.dataframe(st_df.sort_values('Average', ascending=False), hide_index=True, width="stretch",
                 column_config={
                     'Average': st.column_config.NumberColumn(format="%.1f"),
                     'Median': st.column_config.NumberColumn(format="%.1f"),
                     'Highest': st.column_config.NumberColumn(format="%.1f"),
                     'Pass_Rate': st.column_config.ProgressColumn("Pass rate", format="%.0f%%", min_value=0, max_value=100),
                     'Below_Pass': st.column_config.NumberColumn("Below pass"),
                 })

    section_heading("Students below pass in a subject")
    pick = st.selectbox("Subject", st_df['Subject'].tolist(), label_visibility="collapsed", key="subj_pick")
    low = sm[sm[pick] < cut][['ID_clean', pick]].merge(sub[['ID', 'Name', 'Section', 'Contact', 'Att_Pct']],
                                                         left_on='ID_clean', right_on='ID').sort_values(pick)
    if low.empty:
        st.success("Every student in scope passed this subject.")
    else:
        st.dataframe(low[['ID', 'Name', 'Section', 'Contact', pick, 'Att_Pct']], hide_index=True, width="stretch",
                     column_config={pick: st.column_config.NumberColumn("Score", format="%.1f"),
                                    'Att_Pct': st.column_config.NumberColumn("Attendance", format="%.1f%%")})


# ══════════════════════════════════════════════════════════════════════════════
# VIEW: INTERVENTIONS
# ══════════════════════════════════════════════════════════════════════════════
def render_interventions(sub):
    n_drop = int(sub['Dropped_Out'].sum())
    sub = sub[~sub['Dropped_Out']]
    st.markdown("## Intervention list")
    st.caption("Students below the attendance requirement or the pass mark, ordered by urgency."
               + (f" {n_drop} drop-out(s) excluded." if n_drop else ""))
    kind = st.segmented_control("Show", ["All flagged", "Attendance shortfall", "Below pass mark", "Both"],
                                default="All flagged", key="int_kind", label_visibility="collapsed") or "All flagged"
    att_low = sub['Att_Pct'] < att_thr
    mk_low = sub['Marks_Pct'] < pass_thr
    mask = {"All flagged": att_low | mk_low, "Attendance shortfall": att_low,
            "Below pass mark": mk_low, "Both": att_low & mk_low}[kind]
    d = sub[mask].copy()
    if d.empty:
        empty_state("No students flagged", " Nobody in the current scope needs intervention for this criterion.")
        return
    # Urgency: attendance gap + marks gap + failed subjects
    d['Urgency'] = ((att_thr - d['Att_Pct'].fillna(att_thr)).clip(lower=0)
                    + (pass_thr - d['Marks_Pct'].fillna(pass_thr)).clip(lower=0) * 1.5 + d['Failed_Subj'] * 3)
    d = d.sort_values('Urgency', ascending=False)

    pending = d[d['PTM_Status'] == 'Not met']
    kpi_row([
        kpi("Flagged students", f"{len(d)}", f"{len(d) / len(sub) * 100:.0f}% of scope", RED),
        kpi("Critical", f"{(d['Risk'] == 'Critical').sum()}", "highest-priority cases", RED),
        kpi("Avg lectures to recover", f"{d.loc[d['Lec_Needed'] > 0, 'Lec_Needed'].mean():.0f}"
            if (d['Lec_Needed'] > 0).any() else "0", f"to reach {att_thr:.0f}%", AMBER),
        kpi("Not yet consulted", f"{len(pending)}", "PTM status 'Not met'", BLUE),
        kpi("With phone on file", f"{(d['Contact'] != 'N/A').sum()}", "reachable on WhatsApp", GREEN),
    ])

    def msg_for(r):
        return (f"Dear Parent, this is an update from the Academic Mentorship Cell regarding {r['Name']} "
                f"({r['ID']}, {r['Section']}). Current attendance is {fmt_pct(r['Att_Pct'])} against the "
                f"required {att_thr:.0f}%, and aggregate marks are {fmt_pct(r['Marks_Pct'])}. "
                f"Please contact the class mentor at the earliest to discuss an improvement plan.")

    d['WhatsApp'] = d.apply(lambda r: whatsapp_url(r['Contact'], msg_for(r)) if r['Contact'] != 'N/A' else None, axis=1)
    table = d[['ID', 'Name', 'Section', 'Contact', 'Att_Pct', 'Lec_Needed', 'Marks_Pct', 'Failed_Subj',
               'Risk', 'PTM_Status', 'WhatsApp']].reset_index(drop=True)
    st.session_state._int_ids = table['ID'].tolist()

    def _pick():
        rows = st.session_state.int_tbl.selection.rows
        if rows:
            open_student(st.session_state._int_ids[rows[0]])

    with st.container(key="tbl_int"):
        st.dataframe(table, hide_index=True, width="stretch", height=520, key="int_tbl",
                     on_select=_pick, selection_mode="single-row",
                     column_config={
                         'ID': st.column_config.TextColumn("Roll ID", pinned=True),
                         'Name': st.column_config.TextColumn("Name", pinned=True, width="medium"),
                         'Att_Pct': st.column_config.ProgressColumn("Attendance", format="%.1f%%", min_value=0, max_value=100, help="Attendance % counted from the attendance workbook up to the cut-off date."),
                         'Lec_Needed': st.column_config.NumberColumn("Lectures to recover",
                                                                     help=f"Consecutive lectures needed to reach {att_thr:.0f}%"),
                         'Marks_Pct': st.column_config.ProgressColumn("Aggregate", format="%.1f%%", min_value=0, max_value=100, help="Total marks as a percentage of the maximum."),
                         'Failed_Subj': st.column_config.NumberColumn("Subjects < pass", help="Number of subjects scored under the pass mark."),
                         'PTM_Status': st.column_config.TextColumn("PTM", help="Parent-meeting status recorded for this student."),
                         'WhatsApp': st.column_config.LinkColumn("Notify parent", display_text="Open WhatsApp", help="Opens WhatsApp with a ready-made message to the parent."),
                     })
        st.caption("Select a row to open the student's dossier.")

    with st.container(key="cards_int"):
        render_cards(d, "int")

    export = table.drop(columns=['WhatsApp']).rename(columns={
        'Att_Pct': 'Attendance %', 'Lec_Needed': 'Lectures to recover', 'Marks_Pct': 'Aggregate %',
        'Failed_Subj': 'Subjects below pass', 'PTM_Status': 'PTM status'})
    xbuf = io.BytesIO()
    with pd.ExcelWriter(xbuf, engine='openpyxl') as xw:
        export.to_excel(xw, index=False, sheet_name='Interventions')
        for sec, g in export.groupby('Section'):
            g.to_excel(xw, index=False, sheet_name=re.sub(r'[\[\]:*?/\\]', '', str(sec))[:31])
    tag = (sel_sec if sel_sec != "All sections" else "All").replace(' ', '_')
    e1, e2, _ = st.columns([1, 1, 2])
    e1.download_button("Download Excel (by section)", xbuf.getvalue(),
                       file_name=f"PTM_Interventions_{tag}_{datetime.now():%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", icon=":material/table_view:", width="stretch", key="dl_xlsx")
    e2.download_button("Download CSV", export.to_csv(index=False).encode('utf-8'),
                       file_name=f"PTM_Interventions_{tag}_{datetime.now():%Y%m%d}.csv", mime="text/csv",
                       icon=":material/download:", width="stretch", key="dl_csv")


# ══════════════════════════════════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════════════════════════════════
st.write("")
if view == "Overview":
    render_overview(fdf)
elif view == "Student Dossier":
    render_dossier(fdf)
elif view == "Directory":
    render_directory(fdf)
elif view == "Subjects":
    render_subjects(fdf)
elif view == "Interventions":
    render_interventions(fdf)
