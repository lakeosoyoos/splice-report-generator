"""
Splice Report Generator — Streamlit App
========================================
Bidirectional splice QC report from OTDR SOR or JSON files (EXFO FastReporter).

Two-pass analysis:
  Pass 1 — standard bidirectional splice analysis at known splice positions
  Pass 2 — B-direction event scan to catch events the A-direction missed

Launch:  streamlit run app.py
"""

import os
import sys
import tempfile
import io
from contextlib import redirect_stdout

import streamlit as st
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from splicereportmatchexfo import (
    load_all, discover_splices, refine_closure_centers, detect_launch_issues,
    analyze_all, scan_a_standalone_events, scan_b_past_breaks,
    apply_field_gainer_rule,
    build_ribbon_data, write_xlsx,
    REBURN_THRESHOLD, RIBBON_SIZE,
)

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Splice Report Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global styles (Trucordia aesthetic) ───────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap');
@import url('https://fonts.googleapis.com/icon?family=Material+Icons');
@import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200&display=swap');

/* ── Reset & base ────────────────────────────────────────────── */
#MainMenu       { visibility: hidden; }
footer          { visibility: hidden; }
header          { visibility: hidden; }
.stDeployButton { display: none; }

html, body, [class*="css"], .stMarkdown, p, label, span, div {
    font-family: 'Nunito', 'Segoe UI', Arial, sans-serif !important;
}
.main .block-container {
    padding-top: 0rem !important;
    padding-bottom: 2rem;
    max-width: 1200px;
}

/* ── Top utility strip ───────────────────────────────────────── */
.tc-topbar {
    background: #1C2526;
    color: #cccccc;
    font-size: 12px;
    font-weight: 600;
    font-family: 'Nunito', sans-serif;
    padding: 7px 36px;
    display: flex;
    justify-content: flex-end;
    gap: 28px;
    letter-spacing: 0.4px;
}
.tc-topbar span { color: #aaaaaa; cursor: default; }

/* ── Main nav bar ────────────────────────────────────────────── */
.tc-navbar {
    background: #ffffff;
    border-bottom: 2px solid #eeeeee;
    padding: 16px 36px;
    display: flex;
    align-items: center;
    gap: 16px;
}
.tc-logo-icon {
    font-size: 30px;
    font-weight: 900;
    color: #E8461E;
    line-height: 1;
    font-family: 'Nunito', sans-serif;
    transform: skewX(-8deg);
    display: inline-block;
}
.tc-logo-name {
    font-size: 22px;
    font-weight: 900;
    color: #1a1a1a !important;
    font-family: 'Nunito', sans-serif;
    letter-spacing: -0.3px;
}
.tc-navbar-spacer { flex: 1; }
.tc-contact-btn {
    border: 2px solid #E8461E;
    color: #1a1a1a;
    font-family: 'Nunito', sans-serif;
    font-weight: 800;
    font-size: 13px;
    padding: 8px 18px;
    position: relative;
    cursor: default;
    letter-spacing: 0.2px;
}
.tc-contact-btn::after {
    content: "";
    position: absolute;
    bottom: -2px;
    right: -2px;
    width: 10px;
    height: 10px;
    background: #E8461E;
}

/* ── Hero banner ─────────────────────────────────────────────── */
.tc-hero {
    background: linear-gradient(105deg, #E8461E 55%, #c23610 100%);
    padding: 42px 36px 38px 36px;
    margin-bottom: 32px;
}
.tc-hero h1 {
    font-family: 'Nunito', sans-serif;
    font-size: 34px;
    font-weight: 900;
    color: #ffffff;
    margin: 0 0 10px 0;
    line-height: 1.15;
    letter-spacing: -0.3px;
}
.tc-hero p {
    font-family: 'Nunito', sans-serif;
    font-size: 15px;
    color: rgba(255,255,255,0.88);
    margin: 0;
    font-weight: 600;
}

/* ── Section headings ────────────────────────────────────────── */
.tc-section-title {
    font-family: 'Nunito', sans-serif;
    font-size: 22px;
    font-weight: 900;
    color: #1a1a1a;
    margin: 0 0 4px 0;
    letter-spacing: -0.2px;
}
.tc-section-sub {
    font-family: 'Nunito', sans-serif;
    font-size: 14px;
    color: #666;
    font-weight: 600;
    margin: 0 0 20px 0;
}

/* ── Cards ───────────────────────────────────────────────────── */
.tc-card {
    background: #ffffff;
    border: 1px solid #e5e5e5;
    border-top: 4px solid #E8461E;
    border-radius: 4px;
    padding: 24px 26px;
    margin-bottom: 18px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}
.tc-card-title {
    font-family: 'Nunito', sans-serif;
    font-size: 16px;
    font-weight: 900;
    color: #1a1a1a;
    margin: 0 0 14px 0;
    letter-spacing: -0.1px;
}

/* ── Checklist ───────────────────────────────────────────────── */
.tc-list {
    list-style: none;
    padding: 0;
    margin: 0;
}
.tc-list li {
    font-family: 'Nunito', sans-serif;
    font-size: 14px;
    font-weight: 600;
    color: #333;
    padding: 5px 0;
    display: flex;
    align-items: flex-start;
    gap: 10px;
    line-height: 1.5;
}
.tc-list li::before {
    content: "▸";
    color: #E8461E;
    font-weight: 900;
    font-size: 14px;
    flex-shrink: 0;
    margin-top: 1px;
}

/* ── Stat tiles ──────────────────────────────────────────────── */
.tc-stat-row {
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    margin: 0 0 28px 0;
}
.tc-stat {
    background: #fff;
    border: 1px solid #e5e5e5;
    border-top: 4px solid #E8461E;
    border-radius: 4px;
    padding: 16px 20px;
    min-width: 130px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
.tc-stat-label {
    font-family: 'Nunito', sans-serif;
    font-size: 11px;
    font-weight: 800;
    color: #888;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-bottom: 4px;
}
.tc-stat-value {
    font-family: 'Nunito', sans-serif;
    font-size: 30px;
    font-weight: 900;
    color: #1a1a1a;
    line-height: 1;
}
.tc-stat-sub {
    font-family: 'Nunito', sans-serif;
    font-size: 11px;
    font-weight: 600;
    color: #777;
    margin-top: 3px;
}

/* ── Color legend pills ──────────────────────────────────────── */
.tc-legend {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 8px;
}
.tc-pill {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    padding: 5px 12px;
    border-radius: 3px;
    font-family: 'Nunito', sans-serif;
    font-size: 12px;
    font-weight: 700;
    background: #f5f5f5;
    border: 1px solid #ddd;
    color: #333;
}
.tc-swatch {
    width: 12px;
    height: 12px;
    border-radius: 2px;
    flex-shrink: 0;
}

/* ── Sidebar — BLACK BACKGROUND + WHITE TEXT (tech's request) ─── */
[data-testid="stSidebar"],
[data-testid="stSidebar"] *,
[data-testid="stSidebar"] > div,
[data-testid="stSidebar"] section,
[data-testid="stSidebar"] [data-testid="stExpander"],
[data-testid="stSidebar"] details,
[data-testid="stSidebar"] details > summary,
[data-testid="stSidebar"] [data-testid="stExpanderDetails"] {
    background-color: #000000 !important;
}
[data-testid="stSidebar"] {
    border-right: 3px solid #E8461E !important;
    width: 620px !important;
    min-width: 620px !important;
    max-width: 620px !important;
}
[data-testid="stSidebar"] > div:first-child {
    padding-top: 0.5rem;
    width: 620px !important;
}

/* Hide sidebar collapse button */
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"],
button[data-testid="stBaseButton-headerNoPadding"],
[data-testid="stSidebar"] button[title="Collapse sidebar"],
[data-testid="stSidebar"] button[aria-label="Collapse sidebar"] {
    display: none !important;
}

/* All text content — WHITE, unambiguously readable on black */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] h4,
[data-testid="stSidebar"] h5,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div,
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
[data-testid="stSidebar"] [data-testid="stExpander"] summary,
[data-testid="stSidebar"] [data-testid="stExpander"] summary p,
[data-testid="stSidebar"] [data-testid="stExpander"] summary label,
[data-testid="stSidebar"] [data-testid="stExpander"] summary [data-testid="stMarkdownContainer"],
[data-testid="stSidebar"] [data-testid="stExpander"] p,
[data-testid="stSidebar"] [data-testid="stExpander"] label,
[data-testid="stSidebar"] details summary,
[data-testid="stSidebar"] details summary p {
    font-family: 'Nunito', sans-serif !important;
    color: #ffffff !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    text-align: center !important;
    font-weight: 800 !important;
}
[data-testid="stSidebar"] h2 {
    border-bottom: 2px solid #E8461E !important;
    padding-bottom: 4px !important;
}
/* Input fields — black bg, white text, light border */
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stNumberInput input,
[data-testid="stSidebar"] input[type="number"],
[data-testid="stSidebar"] input[type="text"] {
    background: #000000 !important;
    border: 1px solid #aaaaaa !important;
    color: #ffffff !important;
    font-family: 'Nunito', sans-serif !important;
    font-weight: 600 !important;
}
/* Number input +/- buttons */
[data-testid="stSidebar"] .stNumberInput button,
[data-testid="stSidebar"] [data-testid="stNumberInput"] button {
    background: #222222 !important;
    color: #ffffff !important;
    border: 1px solid #aaaaaa !important;
}
[data-testid="stSidebar"] hr {
    border-color: #444 !important;
}
[data-testid="stSidebar"] .stCaption,
[data-testid="stSidebar"] small,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
    color: #cccccc !important;
}
/* Expander summary row (the clickable header of each threshold section) */
[data-testid="stSidebar"] [data-testid="stExpander"] summary,
[data-testid="stSidebar"] details > summary {
    background: #1a1a1a !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    border: 1px solid #444 !important;
    border-radius: 4px !important;
    padding: 8px 10px !important;
}
/* Expander body — the container that holds the inputs */
[data-testid="stSidebar"] [data-testid="stExpanderDetails"],
[data-testid="stSidebar"] details[open] > div {
    background: #000000 !important;
    border: 1px solid #333 !important;
    border-top: none !important;
    padding: 10px !important;
}
/* Checkbox label */
[data-testid="stSidebar"] .stCheckbox label,
[data-testid="stSidebar"] .stCheckbox label * {
    color: #ffffff !important;
}
/* Radio buttons */
[data-testid="stSidebar"] .stRadio label,
[data-testid="stSidebar"] .stRadio label * {
    color: #ffffff !important;
}
/* File-uploader dropzone — give it a visible border on black bg */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"],
[data-testid="stSidebar"] [data-testid="stFileUploader"] section {
    background: #1a1a1a !important;
    border: 2px dashed #E8461E !important;
    border-radius: 6px !important;
    color: #ffffff !important;
    min-height: 80px !important;
    padding: 12px !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] *,
[data-testid="stSidebar"] [data-testid="stFileUploader"] section * {
    color: #ffffff !important;
}
/* Uploaded-file chip (once a file is loaded) */
[data-testid="stSidebar"] [data-testid="stFileUploaderFile"],
[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] * {
    color: #ffffff !important;
    background: #222 !important;
}
/* Fix file uploader — hide instruction text and Add button, keep Upload button only */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] > div {
    display: none !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] {
    justify-content: center !important;
    padding: 8px 0 !important;
}
/* Hide entire dropzone instructions + any Add button once file is loaded */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDeleteBtn"])
[data-testid="stFileUploaderDropzoneInstructions"] {
    display: none !important;
}
/* Hide the small Add button that appears after upload */
[data-testid="stSidebar"] [data-testid="stFileUploader"] button:not([data-testid="stFileUploaderDeleteBtn"]) {
    display: none !important;
}
/* Checkbox label text — transparent background, light text */
[data-testid="stSidebar"] .stCheckbox label {
    background-color: transparent !important;
    background: transparent !important;
}
[data-testid="stSidebar"] .stCheckbox label:hover {
    background-color: transparent !important;
    background: transparent !important;
}
/* The text node inside the label */
[data-testid="stSidebar"] .stCheckbox label > div,
[data-testid="stSidebar"] .stCheckbox label p {
    color: #e8e8e8 !important;
    font-family: 'Nunito', sans-serif !important;
    font-weight: 700 !important;
    background-color: transparent !important;
}
/* Checked box fill — orange, but only the indicator square */
[data-testid="stSidebar"] .stCheckbox [data-baseweb="checkbox"] [aria-checked="true"] > div:first-child {
    background-color: #E8461E !important;
    border-color: #E8461E !important;
}
/* Unchecked box border */
[data-testid="stSidebar"] .stCheckbox [data-baseweb="checkbox"] > div:first-child {
    border-color: #E8461E !important;
}

/* ── Buttons ─────────────────────────────────────────────────── */
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {
    background-color: #E8461E !important;
    border-color: #E8461E !important;
    color: white !important;
    font-family: 'Nunito', sans-serif !important;
    font-weight: 800 !important;
    border-radius: 3px !important;
    letter-spacing: 0.3px;
}
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover {
    background-color: #c23610 !important;
    border-color: #c23610 !important;
}
.stButton > button,
.stDownloadButton > button {
    border-color: #E8461E !important;
    color: #E8461E !important;
    font-family: 'Nunito', sans-serif !important;
    font-weight: 800 !important;
    border-radius: 3px !important;
}
.stButton > button:hover,
.stDownloadButton > button:hover {
    border-color: #c23610 !important;
    color: #c23610 !important;
}

/* ── Progress bar ────────────────────────────────────────────── */
.stProgress > div > div > div > div {
    background-color: #E8461E !important;
}

/* ── Radio ───────────────────────────────────────────────────── */
:root {
    --primary-color: #E8461E !important;
}
/* In the sidebar, render the radio as a chip-style toggle: hide the
   BaseWeb circle entirely (it can't survive the black background
   invariably) and give the whole label a visible border, with an
   orange fill on the checked option. */
[data-testid="stSidebar"] .stRadio [role="radiogroup"] {
    gap: 8px !important;
}
[data-testid="stSidebar"] .stRadio [role="radiogroup"] label {
    font-family: 'Nunito', sans-serif !important;
    font-weight: 700 !important;
    background-color: #1a1a1a !important;
    border: 2px solid #E8461E !important;
    border-radius: 6px !important;
    padding: 6px 14px !important;
    color: #ffffff !important;
    cursor: pointer !important;
}
[data-testid="stSidebar"] .stRadio [role="radiogroup"] label:hover {
    background-color: #2a2a2a !important;
}
[data-testid="stSidebar"] .stRadio [role="radiogroup"] label[data-checked="true"],
[data-testid="stSidebar"] .stRadio [role="radiogroup"] label:has(input:checked) {
    background-color: #E8461E !important;
    color: #ffffff !important;
}
[data-testid="stSidebar"] .stRadio [role="radiogroup"] label * {
    color: #ffffff !important;
}
/* Hide the BaseWeb radio circle entirely in the sidebar — the chip
   style makes it redundant */
[data-testid="stSidebar"] [data-baseweb="radio"] [role="radio"] {
    display: none !important;
}
/* Non-sidebar radios keep their default BaseWeb circle styling */
[data-baseweb="radio"] [role="radio"][aria-checked="true"] div {
    background-color: #E8461E !important;
    border-color: #E8461E !important;
}
[data-baseweb="radio"] [role="radio"] div {
    border-color: #E8461E !important;
}

/* ── Kill the "arrow_drop_down" material-icon text that leaks because
   the global summary * { font-family: Nunito !important } overrides
   the material-icons / material-symbols font, so the ligature resolves
   as literal text.  Hide ANY icon-font element inside expander
   summaries — the fold/unfold visual is obvious from the content
   change alone.  Also restore the icon font for any non-hidden
   icon spans in case Streamlit adds more. ──────────────────── */
[data-testid="stSidebar"] [data-testid="stExpander"] summary span[class*="material"],
[data-testid="stSidebar"] [data-testid="stExpander"] summary i[class*="material"],
[data-testid="stSidebar"] [data-testid="stExpander"] summary span.material-icons,
[data-testid="stSidebar"] [data-testid="stExpander"] summary span.material-symbols-rounded,
[data-testid="stSidebar"] [data-testid="stExpander"] summary span.material-symbols-outlined,
[data-testid="stSidebar"] details > summary span[class*="material"],
[data-testid="stSidebar"] details > summary i[class*="material"] {
    display: none !important;
    visibility: hidden !important;
    width: 0 !important;
    font-size: 0 !important;
}
/* Belt-and-suspenders: any element inside expander summaries with
   data-testid containing 'Icon' (Streamlit's newer pattern) */
[data-testid="stSidebar"] [data-testid="stExpander"] summary [data-testid*="Icon" i],
[data-testid="stSidebar"] [data-testid="stExpander"] summary [data-testid*="icon" i] {
    display: none !important;
}
/* Nuclear: hide any direct child of the summary that isn't a text
   container (Markdown / p / label / svg we explicitly want).  This
   catches any lingering icon element regardless of class/testid. */
[data-testid="stSidebar"] [data-testid="stExpander"] summary > span:not([data-testid="stMarkdownContainer"]):not(:has(p)):not(:has(label)),
[data-testid="stSidebar"] [data-testid="stExpander"] summary > i,
[data-testid="stSidebar"] details > summary > span:not([data-testid="stMarkdownContainer"]):not(:has(p)):not(:has(label)) {
    display: none !important;
    font-size: 0 !important;
    width: 0 !important;
    color: transparent !important;
}

/* ── Equal-height columns ────────────────────────────────────── */
[data-testid="stHorizontalBlock"] {
    align-items: stretch !important;
}
[data-testid="stHorizontalBlock"] > * {
    display: flex !important;
    flex-direction: column !important;
}
[data-testid="stHorizontalBlock"] > * > [data-testid="stVerticalBlock"] {
    flex: 1 !important;
    display: flex !important;
    flex-direction: column !important;
}
[data-testid="stHorizontalBlock"] [data-testid="stMarkdownContainer"] {
    flex: 1 !important;
    display: flex !important;
    flex-direction: column !important;
}
.tc-card {
    flex: 1 !important;
    margin-bottom: 0 !important;
}

/* ── Links ───────────────────────────────────────────────────── */
a { color: #E8461E !important; }

/* ── Expander ────────────────────────────────────────────────── */
[data-testid="stExpander"] {
    border: 1px solid #e0e0e0 !important;
    border-radius: 4px !important;
    background: #ffffff !important;
}
[data-testid="stExpander"] summary {
    font-family: 'Nunito', sans-serif !important;
    font-weight: 800 !important;
    color: #1a1a1a !important;
    background: #ffffff !important;
}
[data-testid="stExpander"] summary:hover {
    color: #E8461E !important;
}
[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary span,
[data-testid="stExpander"] summary * {
    color: #1a1a1a !important;
    font-family: 'Nunito', sans-serif !important;
    font-weight: 800 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Navigation ────────────────────────────────────────────────────────────────

st.markdown("""
<div class="tc-topbar">
    <span>Splice QC Tools</span>
    <span>Help</span>
</div>
<div class="tc-navbar">
    <div class="tc-logo-icon">↗</div>
    <a href="/" target="_self" class="tc-logo-name" style="text-decoration:none; cursor:pointer;">Splice Report Generator</a>
    <div class="tc-navbar-spacer"></div>
    <div class="tc-contact-btn">OTDR QC &nbsp; ▸</div>
</div>
""", unsafe_allow_html=True)


# ── Password protection ───────────────────────────────────────────────────────

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    try:
        correct = st.secrets["passwords"]["app_password"]
    except (KeyError, FileNotFoundError):
        return True
    pwd = st.text_input("Enter password", type="password", key="pwd_input")
    if pwd:
        if pwd == correct:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password")
    return False

if not check_password():
    st.stop()


# ── Session state ─────────────────────────────────────────────────────────────

for key in ["xlsx_bytes", "xlsx_name", "summary_data", "log_output", "done",
            "zach_pdf_bytes", "zach_pdf_name"]:
    if key not in st.session_state:
        st.session_state[key] = None
if "upload_key" not in st.session_state:
    st.session_state.upload_key = 0
if "downloaded_once" not in st.session_state:
    st.session_state.downloaded_once = False


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## Upload SOR Files")

    # Single uploader per direction — accepts loose .sor / .json files OR
    # a .zip archive of them.  The run block auto-detects based on the
    # actual files that got dropped in.
    uploaded_a = st.file_uploader(
        "A-direction files (.sor, .json, or .zip)",
        type=["sor", "json", "zip"],
        accept_multiple_files=True,
        key=f"upload_a_{st.session_state.upload_key}",
    )
    if uploaded_a:
        total_kb = sum(f.size for f in uploaded_a) / 1024
        st.caption(f"A: {len(uploaded_a)} file(s), {total_kb:.0f} KB total")
    uploaded_b = st.file_uploader(
        "B-direction files (.sor, .json, or .zip)",
        type=["sor", "json", "zip"],
        accept_multiple_files=True,
        key=f"upload_b_{st.session_state.upload_key}",
    )
    if uploaded_b:
        total_kb = sum(f.size for f in uploaded_b) / 1024
        st.caption(f"B: {len(uploaded_b)} file(s), {total_kb:.0f} KB total")
    if st.button("Clear All", use_container_width=True):
        old_key = st.session_state.upload_key
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.session_state.upload_key = old_key + 1
        st.rerun()

    # ── Settings locked to the flowchart-PDF defaults ────────────────
    # All threshold controls were removed per tech direction.  The
    # script runs with the gate set documented in
    # SCRIPT_LOGIC_FLOWCHART.pdf (April 28 revision):
    #   • Reburn threshold 0.160 dB (pink A+B reburn)
    #   • Bend: positive loss ≥ 0.090 dB AND offset > 150 m from per-
    #     fiber splice km AND per-fiber length-model residual ≥ 150 m
    #     AND narrow-LSA at predicted km shows real loss step (yellow)
    #   • Field gainer: bidir avg in [-0.7, 0] dB AND both A+B real
    #     measurements with opposite signs (mint)
    #   • In-line REF: reflective + Fresnel + trace continues past
    #     event (deep orange E64A19) — formerly tagged as BREAK
    #   • Closure validation: loss-distribution gate only
    #     (gainer_frac < 0.05 AND median_loss > 0.100 dB → phantom)
    #   • Launch: reflectance > -50 dB (orange; loss rule off)
    #   • Dead-zone annotation for broken fibers where B also ends short
    threshold   = REBURN_THRESHOLD
    ribbon_size = RIBBON_SIZE

    st.divider()
    st.markdown(
        "<div style='color:#ccc;font-size:11.5px;line-height:1.45;padding:4px 0;'>"
        "Running with the current default gates (see "
        "<b>SCRIPT_LOGIC_FLOWCHART.pdf</b>).  Thresholds are not "
        "adjustable in the UI — the script is tuned against the tech's "
        "reference reports."
        "</div>",
        unsafe_allow_html=True,
    )

    has_a = bool(uploaded_a)
    has_b = bool(uploaded_b)
    run_button = st.button("Generate Report", type="primary",
                           use_container_width=True,
                           disabled=not (has_a and has_b))
    if has_a and not has_b:
        st.caption(
            ":warning: B-direction files required.  A splice report needs "
            "both directions to compute the bidirectional average — "
            "single-direction OTDR readings can be biased by gainers, "
            "fiber-lot mismatches, and noise that the bidirectional average "
            "cancels out.")


# ── Helpers ───────────────────────────────────────────────────────────────────

def detect_sites(dir_a, dir_b=None):
    """
    Auto-detect site A and B names from the folder name or first SOR/JSON filename.
    e.g. folder 'TULBAR' → ('TUL', 'BAR'), file 'STRROM001.sor' → ('STR', 'ROM')
    Falls back to ('A', 'B') if nothing parseable is found.
    """
    candidates = []
    # Folder name of dir_a
    candidates.append(os.path.basename(os.path.normpath(dir_a)).upper())
    # First SOR or JSON filename in dir_a
    try:
        trace_files = sorted([f for f in os.listdir(dir_a)
                              if f.lower().endswith('.sor') or f.lower().endswith('.json')])
        if trace_files:
            # Strip leading alphabetic prefix (e.g. "NEWELM001 .json" → "NEWELM")
            name = trace_files[0].split('.')[0].split('_')[0].split(' ')[0].upper()
            candidates.append(name)
    except Exception:
        pass

    for name in candidates:
        # Strip common suffixes
        for suffix in ['_1550', '_1310', '_SOR', '_FILES']:
            name = name.replace(suffix, '')
        alpha = ''.join(c for c in name if c.isalpha())
        # Prefer clean 6-char split: first 3 = site A, next 3 = site B
        if len(alpha) == 6:
            return alpha[:3], alpha[3:]
        # Allow 3+4 or 4+3
        if len(alpha) in (7, 8):
            mid = len(alpha) // 2
            return alpha[:mid], alpha[mid:]

    return 'A', 'B'


def stage_files(uploaded, prefix="trace_"):
    """Stage uploaded SOR or JSON files into a temp directory.  The script's
    loader auto-detects file type per directory."""
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    for uf in uploaded:
        with open(os.path.join(tmpdir, uf.name), 'wb') as f:
            f.write(uf.getbuffer())
    return tmpdir


def stage_inputs(uploaded_list, prefix="trace_"):
    """Stage a mixed list of uploads (.sor / .json / .zip).  Each file is
    routed by extension — ZIPs are extracted in place, loose OTDR files
    are copied directly into the same temp dir.  Returns (tmpdir, n, errors)
    where n is the total OTDR-file count landed and errors is a list of
    friendly strings for any unreadable archive."""
    import zipfile
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    n = 0
    errors = []
    for uf in uploaded_list:
        name = uf.name
        lower = name.lower()
        if lower.endswith('.zip'):
            try:
                uf.seek(0)
            except Exception:
                pass
            raw = uf.read()
            try:
                with zipfile.ZipFile(io.BytesIO(raw), 'r') as zf:
                    for zname in zf.namelist():
                        zlower = zname.lower()
                        if not (zlower.endswith('.sor') or zlower.endswith('.json')):
                            continue
                        if zname.startswith('__MACOSX') or '/.DS_Store' in zname:
                            continue
                        basename = os.path.basename(zname)
                        if not basename:
                            continue
                        with zf.open(zname) as src, open(os.path.join(tmpdir, basename), 'wb') as dst:
                            dst.write(src.read())
                        n += 1
            except zipfile.BadZipFile as e:
                errors.append(
                    f"Could not read '{name}' as a ZIP ({e}).  "
                    f"Re-zip the folder of .sor/.json files."
                )
        elif lower.endswith('.sor') or lower.endswith('.json'):
            try:
                uf.seek(0)
            except Exception:
                pass
            with open(os.path.join(tmpdir, name), 'wb') as f:
                f.write(uf.read())
            n += 1
        # Silently skip any other file extension — st.file_uploader
        # already restricts to .sor/.json/.zip.
    return tmpdir, n, errors


def stage_zip(uploaded_zip, prefix="trace_zip_"):
    """Extract SOR and/or JSON files from an uploaded ZIP into a temp dir.

    Returns (tmpdir, n_extracted).  Caller checks n_extracted > 0 and
    surfaces a friendly error if it's zero.

    Implementation notes:
    - Always rewinds the upload buffer first; Streamlit's UploadedFile
      may have its seek pointer at EOF if it was previously read.
    - Reads the bytes once with .read() and wraps in BytesIO — more
      robust than .getbuffer() (which returns a memoryview that some
      Python / zipfile combinations have trouble with on cloud).
    """
    import zipfile
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    try:
        uploaded_zip.seek(0)
    except Exception:
        pass
    raw = uploaded_zip.read()
    n_extracted = 0
    try:
        with zipfile.ZipFile(io.BytesIO(raw), 'r') as zf:
            for name in zf.namelist():
                lower = name.lower()
                if not (lower.endswith('.sor') or lower.endswith('.json')):
                    continue
                if name.startswith('__MACOSX') or '/.DS_Store' in name:
                    continue
                basename = os.path.basename(name)
                if not basename:
                    continue
                with zf.open(name) as src, open(os.path.join(tmpdir, basename), 'wb') as dst:
                    dst.write(src.read())
                n_extracted += 1
    except zipfile.BadZipFile as e:
        raise RuntimeError(
            f"Could not read '{getattr(uploaded_zip, 'name', '?')}' as a ZIP "
            f"file ({e}).  Re-zip the folder of .sor/.json files and try again."
        )
    return tmpdir, n_extracted


# ── Run ───────────────────────────────────────────────────────────────────────

if run_button and has_a:
    prog = st.progress(0.0, text="Staging A-direction input...")
    dir_a, n_a, errs_a = stage_inputs(uploaded_a, "splice_a_")
    for e in errs_a:
        st.warning(e)
    if n_a == 0:
        prog.empty()
        st.error(
            "No .sor or .json files found in the A-direction input.  Drop "
            "either loose OTDR files or a ZIP containing them."
        )
        st.stop()

    prog.progress(0.4, text=f"A staged ({n_a} files).  Staging B-direction...")
    dir_b = None
    if uploaded_b:
        dir_b, n_b, errs_b = stage_inputs(uploaded_b, "splice_b_")
        for e in errs_b:
            st.warning(e)
        if n_b == 0:
            prog.empty()
            st.error("No .sor or .json files found in the B-direction input.")
            st.stop()
    prog.progress(0.5, text="Files staged.")
    prog.empty()

    # Auto-detect site names from folder/filenames
    site_a, site_b = detect_sites(dir_a, dir_b)

    bar     = st.progress(0.0, text="Loading OTDR files (SOR or JSON)...")
    log_buf = io.StringIO()

    with redirect_stdout(log_buf):
        fibers_a, fibers_b = load_all(dir_a, dir_b)
    n_fibers = max(fibers_a.keys()) if fibers_a else 0
    bar.progress(0.15, text=f"Loaded {len(fibers_a)} A + {len(fibers_b)} B fibers...")

    with redirect_stdout(log_buf):
        splice_candidates = discover_splices(fibers_a)
        real_splices, phantom_zones = refine_closure_centers(
            fibers_a, splice_candidates, return_phantoms=True)
        # Interleave phantom bend/damage zones between the real splices in
        # position order — mirrors the tech's Cle Elum layout.
        splices = sorted(
            list(real_splices) + list(phantom_zones),
            key=lambda sp: sp.get('position_km_refined', sp['position_km']),
        )
        splice_display_num = 0
        for sp in splices:
            if sp.get('column_kind') == 'splice':
                splice_display_num += 1
                sp['splice_display_num'] = splice_display_num
    bar.progress(0.25, text=f"Found {len(splices)} splice closures...")

    # Launch-issue detection (module defaults — mirrors the CLI)
    bar.progress(0.30, text="Detecting launch-end issues...")
    first_splice_km = splices[0]['position_km'] if splices else None
    with redirect_stdout(log_buf):
        launch_issues = detect_launch_issues(fibers_a, fibers_b, first_splice_km)

    actual_span = 0
    if actual_span == 0:
        all_ends = sorted([e['dist_km'] for r in fibers_a.values()
                           for e in r['events'] if e['is_end']])
        if all_ends:
            top_q = all_ends[int(len(all_ends) * 0.75):]
            actual_span = round(np.median(top_q), 2)

    # Pass 1 — analyze_all (at-splice classification + broke + B-fill)
    bar.progress(0.45, text=f"Pass 1: {n_fibers} fibers × {len(splices)} splice positions...")
    with redirect_stdout(log_buf):
        results = analyze_all(fibers_a, fibers_b, splices, threshold)

    # Pass 2a — standalone A-direction bends / breaks not at a closure
    bar.progress(0.65, text="Pass 2a: scanning A-direction standalone events...")
    with redirect_stdout(log_buf):
        a_standalone = scan_a_standalone_events(
            fibers_a, splices, results, actual_span,
        )

    # Pass 2b — past-break B-fill (B-direction only used past A-side breaks)
    bar.progress(0.72, text="Pass 2b: scanning B past A-side breaks (B-fill)...")
    with redirect_stdout(log_buf):
        b_pastbreak = scan_b_past_breaks(
            fibers_a, fibers_b, splices, threshold, results, actual_span,
        )

    # Merge in the CLI's priority order: Pass 1 > Pass 2a > Pass 2b
    all_results = {**results, **a_standalone, **b_pastbreak}

    # Field-gainer post-pass — flag mid-span events with loss in [-0.7, 0]
    with redirect_stdout(log_buf):
        apply_field_gainer_rule(all_results, actual_span)

    n_reburn      = sum(1 for r in all_results.values()
                        if r.get('event_source') in ('bidir', 'bidir_grey_a', 'bidir_grey_b')
                        and not r['is_break'] and not r.get('is_bend'))
    n_breaks      = sum(1 for r in all_results.values() if r['is_break'])
    n_broke       = sum(1 for r in all_results.values() if r['is_broke'])
    n_bfill       = sum(1 for r in all_results.values() if r.get('is_bfill', False))
    n_bend        = sum(1 for r in all_results.values() if r.get('is_bend', False))
    n_bend_high   = sum(1 for r in all_results.values()
                        if r.get('is_bend') and r.get('bend_severity') == 'HIGH')
    n_bend_review = sum(1 for r in all_results.values()
                        if r.get('is_bend') and r.get('bend_severity') == 'REVIEW')
    n_bend_watch  = sum(1 for r in all_results.values()
                        if r.get('is_bend') and r.get('bend_severity') == 'WATCH')
    n_a_only      = sum(1 for r in all_results.values() if r.get('is_a_only', False))
    n_b_only      = sum(1 for r in all_results.values() if r.get('is_b_only', False))
    n_b_only_high = sum(1 for r in all_results.values() if r.get('is_b_only') and r.get('est_bidir_flagged'))
    n_a_only_high = sum(1 for r in all_results.values() if r.get('is_a_only') and r.get('est_bidir_flagged'))
    n_launch      = len(launch_issues)
    n_launch_high = sum(1 for v in launch_issues.values() if v.get('severity') == 'HIGH')

    bar.progress(0.80, text="Building ribbon grid...")
    with redirect_stdout(log_buf):
        cells, launch_cells_a, launch_cells_b = build_ribbon_data(
            all_results, n_fibers, ribbon_size, len(splices),
            launch_issues=launch_issues,
        )

    bar.progress(0.92, text="Writing Excel report...")
    xlsx_dir  = tempfile.mkdtemp(prefix="splice_xlsx_")
    xlsx_path = os.path.join(xlsx_dir, "splice_report.xlsx")
    with redirect_stdout(log_buf):
        write_xlsx(cells, splices, n_fibers, ribbon_size, xlsx_path,
                   site_a, site_b, actual_span,
                   launch_cells_a=launch_cells_a, launch_cells_b=launch_cells_b)

    with open(xlsx_path, 'rb') as f:
        st.session_state.xlsx_bytes = f.read()
    st.session_state.xlsx_name   = f"splice_report_{site_a}_{site_b}.xlsx"

    # ── Build Zach's Explanation PDF (per-cell explanation, no tech compare) ──
    try:
        from make_zach_explanation import build_explanation_pdf
        zach_path = os.path.join(xlsx_dir, "zach_explanation.pdf")
        with redirect_stdout(log_buf):
            build_explanation_pdf(
                all_results, splices, launch_issues, actual_span,
                site_a, site_b, zach_path,
                ribbon_size=ribbon_size,
                reburn_threshold=threshold,
            )
        with open(zach_path, 'rb') as f:
            st.session_state.zach_pdf_bytes = f.read()
        st.session_state.zach_pdf_name = (
            f"zach_explanation_{site_a}_{site_b}.pdf")
    except Exception as exc:
        # Don't let a PDF-build failure block the xlsx delivery.
        st.session_state.zach_pdf_bytes = None
        st.session_state.zach_pdf_name = None
        log_buf.write(f"\n[zach-explanation] failed: {exc}\n")
    st.session_state.summary_data = dict(
        n_fibers=n_fibers, n_splices=len(splices), actual_span=actual_span,
        threshold=threshold, n_flagged=len(all_results),
        n_reburn=n_reburn, n_breaks=n_breaks, n_broke=n_broke,
        n_bfill=n_bfill, n_a_only=n_a_only, n_a_only_high=n_a_only_high,
        n_b_only=n_b_only, n_b_only_high=n_b_only_high,
        n_bend=n_bend, n_bend_high=n_bend_high,
        n_bend_review=n_bend_review, n_bend_watch=n_bend_watch,
        n_launch=n_launch, n_launch_high=n_launch_high,
        site_a=site_a, site_b=site_b,
    )
    st.session_state.log_output = log_buf.getvalue()
    st.session_state.done       = True
    # Reset so the auto-download triggers on this fresh report
    st.session_state.downloaded_once = False

    bar.progress(1.0, text="Done!")
    bar.empty()


# ── Display ───────────────────────────────────────────────────────────────────

if st.session_state.get("done"):
    d = st.session_state.summary_data

    # Compact "Report ready" banner — minimal, just the direction + key numbers
    st.markdown(f"""
    <div class="tc-hero" style="padding-top:24px; padding-bottom:24px;">
        <h1>Report Ready&nbsp;·&nbsp;{d['site_a']} → {d['site_b']}</h1>
        <p>{d['n_fibers']} fibers &nbsp;·&nbsp; {d['n_flagged']} flagged events
           &nbsp;·&nbsp; Threshold: {d['threshold']:.3f} dB</p>
    </div>
    """, unsafe_allow_html=True)

    # Auto-download the Excel file AND Zach's Explanation PDF on first render
    # after the report completes.  Manual fallback download buttons for both
    # in case the browser blocks the auto-download.
    if st.session_state.xlsx_bytes:
        import base64
        import streamlit.components.v1 as components

        just_generated = not st.session_state.get("downloaded_once", False)
        if just_generated:
            xlsx_b64 = base64.b64encode(st.session_state.xlsx_bytes).decode()
            zach_bytes = st.session_state.get("zach_pdf_bytes")
            zach_name  = st.session_state.get("zach_pdf_name")
            zach_b64   = (base64.b64encode(zach_bytes).decode()
                           if zach_bytes else None)
            # Build a 2-link auto-download component.  Stagger the second
            # click by ~600 ms so the browser treats it as a separate
            # user-initiated download rather than collapsing both into one.
            zach_block = ""
            if zach_b64 and zach_name:
                zach_block = f"""
            <a id="auto_dl_pdf"
               href="data:application/pdf;base64,{zach_b64}"
               download="{zach_name}"></a>"""
            zach_click = (
                "setTimeout(() => {"
                "  const p = document.getElementById('auto_dl_pdf');"
                "  if (p) p.click();"
                "}, 600);" if zach_b64 else ""
            )
            components.html(f"""
            <html><body>
            <a id="auto_dl"
               href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{xlsx_b64}"
               download="{st.session_state.xlsx_name}"></a>{zach_block}
            <script>
              const xlsx = document.getElementById('auto_dl');
              if (xlsx) {{ xlsx.click(); }}
              {zach_click}
            </script>
            </body></html>
            """, height=0)
            st.session_state.downloaded_once = True

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            "⬇  Download Excel Again",
            st.session_state.xlsx_bytes,
            file_name=st.session_state.xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        zach_bytes = st.session_state.get("zach_pdf_bytes")
        zach_name  = st.session_state.get("zach_pdf_name")
        if zach_bytes:
            st.download_button(
                "⬇  Download Zach's Explanation Again",
                zach_bytes,
                file_name=zach_name or "zach_explanation.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        st.caption(
            "Both files download automatically when ready: the Excel report "
            "and Zach's Explanation PDF.  If your browser blocked either "
            "auto-download, use the buttons above."
        )



else:
    # ── Landing ───────────────────────────────────────────────────────────────

    st.markdown("""
    <div class="tc-hero">
        <h1>Where Accuracy<br>Meets Every Fiber</h1>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <p class="tc-section-title" style="color:#ffffff;">Splice Quality Customized to Your Span</p>
    <p class="tc-section-sub" style="color:#ffffff;">Upload your A and B direction SOR files. The report finds every flagged event - whether one direction saw it or both.</p>
    """, unsafe_allow_html=True)

    # ── Top row: Pass 1 / Pass 2 ──────────────────────────────────────────────
    st.markdown("""
    <div style="display:flex; gap:18px; align-items:stretch; margin-bottom:18px;">
        <div class="tc-card" style="flex:1; margin-bottom:0;">
            <div class="tc-card-title">Pass 1 - Splice Position Analysis</div>
            <ul class="tc-list">
                <li>Discovers splice closure positions where 20+ fibers share an event</li>
                <li>Finds A+B bidirectional events and flags if loss >= threshold</li>
                <li>Detects broke fibers (mid-span trace termination)</li>
                <li>Fills B-direction data past breaks where A is blind</li>
                <li>Flags A-only events with estimated bidir = A / 2</li>
            </ul>
        </div>
        <div class="tc-card" style="flex:1; margin-bottom:0;">
            <div class="tc-card-title">Pass 2 - B-Direction Event Scan</div>
            <ul class="tc-list">
                <li>Scans every B-direction event above threshold not caught in Pass 1</li>
                <li>Converts B-frame positions to A-frame coordinates</li>
                <li>Matches to nearest splice position within 1.5 km</li>
                <li>If A event also found: computes true bidirectional average</li>
                <li>If no A event: flags as B-only with estimated bidir = B / 2</li>
                <li>Catches events regardless of which direction saw it first</li>
            </ul>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Bottom row: How To Use / Color Key ────────────────────────────────────
    st.markdown("""
    <div style="display:flex; gap:18px; align-items:stretch; margin-bottom:18px;">
        <div class="tc-card" style="flex:1; margin-bottom:0;">
            <div class="tc-card-title">How To Use</div>
            <ul class="tc-list">
                <li>Upload A-direction and B-direction SOR/JSON files (both required) as a ZIP or individual files</li>
                <li>Adjust the reburn threshold if needed - default is 0.160 dB</li>
                <li>Use the <strong>Include in Report</strong> checkboxes to filter which event types appear in the output</li>
                <li>Click <strong>Generate Report</strong> to download the color-coded Excel splice QC report</li>
            </ul>
        </div>
        <div class="tc-card" style="flex:1; margin-bottom:0;">
            <div class="tc-card-title">Excel Report Color Key</div>
            <div class="tc-legend">
                <span class="tc-pill"><span class="tc-swatch" style="background:#FFC7CE"></span>Pink - A+B Reburn</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FF4444"></span>Red - Break / Broke</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#BDD7EE"></span>Blue - B-fill past break</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#BFBFBF"></span>Gray - Dead zone</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FFEB3B"></span>Yellow - Bend (&ge; 0.090 dB)</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FFF2CC"></span>Lt. Yellow - A-only OK</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FF7043"></span>Coral - A-only &#9888;</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#E8D5F5"></span>Lavender - B-only OK</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#C084FC"></span>Purple - B-only &#9888;</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FFA500"></span>Orange - Launch issue</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
