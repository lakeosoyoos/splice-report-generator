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
    load_all, discover_splices, analyze_all, scan_b_events,
    build_ribbon_data, write_xlsx,
    REBURN_THRESHOLD, NOMINAL_SPLICE, RIBBON_SIZE,
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

/* ── Sidebar ─────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background-color: #1C2526 !important;
    border-right: 3px solid #E8461E !important;
    width: 620px !important;
    min-width: 620px !important;
    max-width: 620px !important;
}
[data-testid="stSidebar"] > div:first-child {
    padding-top: 0.5rem;
    width: 620px !important;
}

/* ── Hide sidebar collapse button (removes keyboard double) ──── */
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"],
button[data-testid="stBaseButton-headerNoPadding"],
[data-testid="stSidebar"] button[title="Collapse sidebar"],
[data-testid="stSidebar"] button[aria-label="Collapse sidebar"] {
    display: none !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
    font-family: 'Nunito', sans-serif !important;
    color: #e8e8e8 !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    text-align: center !important;
}
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stNumberInput input {
    background: #243030 !important;
    border-color: #E8461E !important;
    color: #e8e8e8 !important;
    font-family: 'Nunito', sans-serif !important;
}
[data-testid="stSidebar"] hr {
    border-color: #2e3d3d !important;
}
[data-testid="stSidebar"] .stCaption,
[data-testid="stSidebar"] small {
    color: #aaa !important;
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
.stRadio [role="radiogroup"] label {
    font-family: 'Nunito', sans-serif !important;
    font-weight: 700 !important;
    background-color: transparent !important;
    border-color: transparent !important;
}
.stRadio [role="radiogroup"] label[data-checked="true"],
.stRadio [role="radiogroup"] label:has(input:checked) {
    background-color: transparent !important;
    border-color: transparent !important;
    color: inherit !important;
}
.stRadio [role="radiogroup"] label[data-checked="true"] p,
.stRadio [role="radiogroup"] label:has(input:checked) p {
    color: inherit !important;
}
/* Override Streamlit/BaseWeb primary color for radio dot */
:root {
    --primary-color: #E8461E !important;
}
[data-baseweb="radio"] [role="radio"][aria-checked="true"] div {
    background-color: #E8461E !important;
    border-color: #E8461E !important;
}
[data-baseweb="radio"] [role="radio"] div {
    border-color: #E8461E !important;
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

for key in ["xlsx_bytes", "xlsx_name", "summary_data", "log_output", "done"]:
    if key not in st.session_state:
        st.session_state[key] = None
if "upload_key" not in st.session_state:
    st.session_state.upload_key = 0
if "downloaded_once" not in st.session_state:
    st.session_state.downloaded_once = False


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## Upload SOR Files")

    input_method = st.radio(
        "Input method",
        ["Upload ZIP", "Browse files"],
        index=0,
        horizontal=True,
    )

    uploaded_a = None
    uploaded_b = None
    zip_a      = None
    zip_b      = None

    if input_method == "Upload ZIP":
        zip_a = st.file_uploader("A-direction ZIP", type=["zip"],
                                 accept_multiple_files=False,
                                 key=f"zip_a_{st.session_state.upload_key}")
        if zip_a:
            st.caption(f"A: {zip_a.name} ({zip_a.size/1024:.0f} KB)")
        zip_b = st.file_uploader("B-direction ZIP (optional)", type=["zip"],
                                 accept_multiple_files=False,
                                 key=f"zip_b_{st.session_state.upload_key}")
        if zip_b:
            st.caption(f"B: {zip_b.name} ({zip_b.size/1024:.0f} KB)")

    else:
        uploaded_a = st.file_uploader("A-direction files (.sor or .json)",
                                      type=["sor", "json"],
                                      accept_multiple_files=True,
                                      key=f"upload_a_{st.session_state.upload_key}")
        uploaded_b = st.file_uploader("B-direction files (.sor or .json, optional)",
                                      type=["sor", "json"],
                                      accept_multiple_files=True,
                                      key=f"upload_b_{st.session_state.upload_key}")
    if st.button("Clear All", use_container_width=True):
        old_key = st.session_state.upload_key
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.session_state.upload_key = old_key + 1
        st.rerun()

    st.divider()
    st.markdown("## Settings")

    threshold   = st.number_input("Bidirectional threshold (dB, 0.15=auto)", value=REBURN_THRESHOLD,
                                  format="%.3f", step=0.01)
    ribbon_size = RIBBON_SIZE

    st.markdown("**Include in Report**")
    col_chk1, col_chk2 = st.columns(2)
    with col_chk1:
        inc_reburn = st.checkbox("A+B Reburn", value=True, key="inc_reburn")
        inc_break  = st.checkbox("Break",      value=True, key="inc_break")
        inc_broke  = st.checkbox("Broke",      value=True, key="inc_broke")
    with col_chk2:
        inc_bfill  = st.checkbox("B-fill",     value=True, key="inc_bfill")
        inc_a_only = st.checkbox("A-only",     value=True, key="inc_a_only")
        inc_b_only = st.checkbox("B-only",     value=True, key="inc_b_only")

    has_a = (bool(uploaded_a) or bool(zip_a))
    run_button = st.button("Generate Report", type="primary",
                           use_container_width=True, disabled=not has_a)


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


def stage_zip(uploaded_zip, prefix="trace_zip_"):
    """Extract SOR and/or JSON files from an uploaded ZIP into a temp dir."""
    import zipfile
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    with zipfile.ZipFile(io.BytesIO(uploaded_zip.getbuffer()), 'r') as zf:
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
    return tmpdir


# ── Run ───────────────────────────────────────────────────────────────────────

if run_button and has_a:
    if zip_a:
        prog = st.progress(0.0, text="Extracting A-direction ZIP...")
        dir_a = stage_zip(zip_a, "splice_a_")
        prog.progress(0.4, text="Extracting B-direction ZIP...")
        dir_b = stage_zip(zip_b, "splice_b_") if zip_b else None
        prog.progress(0.5, text="Files extracted.")
        prog.empty()
    else:
        prog = st.progress(0.0, text="Staging A-direction files...")
        dir_a = stage_files(uploaded_a, "splice_a_")
        prog.progress(0.4, text="Staging B-direction files...")
        dir_b = stage_files(uploaded_b, "splice_b_") if uploaded_b else None
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
        splices = discover_splices(fibers_a)
    bar.progress(0.30, text=f"Found {len(splices)} splice closures...")

    actual_span = 0
    if actual_span == 0:
        all_ends = sorted([e['dist_km'] for r in fibers_a.values()
                           for e in r['events'] if e['is_end']])
        if all_ends:
            top_q = all_ends[int(len(all_ends) * 0.75):]
            actual_span = round(np.median(top_q), 2)

    bar.progress(0.45, text=f"Pass 1: {n_fibers} fibers × {len(splices)} splice positions...")
    with redirect_stdout(log_buf):
        results = analyze_all(fibers_a, fibers_b, splices, threshold)

    bar.progress(0.65, text="Pass 2: scanning B-direction for missed events...")
    with redirect_stdout(log_buf):
        b_results = scan_b_events(fibers_a, fibers_b, splices, threshold, results, actual_span)

    all_results = {**results, **b_results}

    # ── Apply event-type filters from sidebar ──────────────────────────────
    def _included(r):
        if r.get('is_break')  and not st.session_state.get('inc_break',  True): return False
        if r.get('is_broke')  and not st.session_state.get('inc_broke',  True): return False
        if r.get('is_bfill')  and not st.session_state.get('inc_bfill',  True): return False
        if r.get('is_a_only') and not st.session_state.get('inc_a_only', True): return False
        if r.get('is_b_only') and not st.session_state.get('inc_b_only', True): return False
        # bidir reburn = not break, not broke, not bfill, not a_only, not b_only
        is_reburn = (r.get('event_source') == 'bidir'
                     and not r.get('is_break') and not r.get('is_broke')
                     and not r.get('is_bfill') and not r.get('is_a_only')
                     and not r.get('is_b_only'))
        if is_reburn and not st.session_state.get('inc_reburn', True): return False
        return True
    all_results = {k: v for k, v in all_results.items() if _included(v)}

    n_reburn      = sum(1 for r in all_results.values() if r.get('event_source') == 'bidir' and not r['is_break'])
    n_breaks      = sum(1 for r in all_results.values() if r['is_break'])
    n_broke       = sum(1 for r in all_results.values() if r['is_broke'])
    n_bfill       = sum(1 for r in all_results.values() if r.get('is_bfill', False))
    n_a_only      = sum(1 for r in all_results.values() if r.get('is_a_only', False))
    n_b_only      = sum(1 for r in all_results.values() if r.get('is_b_only', False))
    n_b_only_high = sum(1 for r in all_results.values() if r.get('is_b_only') and r.get('est_bidir_flagged'))
    n_a_only_high = sum(1 for r in all_results.values() if r.get('is_a_only') and r.get('est_bidir_flagged'))

    bar.progress(0.80, text="Building ribbon grid...")
    with redirect_stdout(log_buf):
        cells = build_ribbon_data(all_results, n_fibers, ribbon_size, len(splices))

    bar.progress(0.92, text="Writing Excel report...")
    xlsx_dir  = tempfile.mkdtemp(prefix="splice_xlsx_")
    xlsx_path = os.path.join(xlsx_dir, "splice_report.xlsx")
    with redirect_stdout(log_buf):
        write_xlsx(cells, splices, n_fibers, ribbon_size, xlsx_path,
                   site_a, site_b, actual_span)

    with open(xlsx_path, 'rb') as f:
        st.session_state.xlsx_bytes = f.read()
    st.session_state.xlsx_name   = f"splice_report_{site_a}_{site_b}.xlsx"
    st.session_state.summary_data = dict(
        n_fibers=n_fibers, n_splices=len(splices), actual_span=actual_span,
        threshold=threshold, n_flagged=len(all_results),
        n_reburn=n_reburn, n_breaks=n_breaks, n_broke=n_broke,
        n_bfill=n_bfill, n_a_only=n_a_only, n_a_only_high=n_a_only_high,
        n_b_only=n_b_only, n_b_only_high=n_b_only_high,
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

    # Auto-download the Excel file on first render after the report completes,
    # then fall back to a manual download button in case the browser blocked
    # the auto-download.
    if st.session_state.xlsx_bytes:
        import base64
        import streamlit.components.v1 as components

        just_generated = not st.session_state.get("downloaded_once", False)
        if just_generated:
            b64 = base64.b64encode(st.session_state.xlsx_bytes).decode()
            components.html(f"""
            <html><body>
            <a id="auto_dl"
               href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
               download="{st.session_state.xlsx_name}"></a>
            <script>
              // Trigger the download as soon as the component mounts.
              const link = document.getElementById('auto_dl');
              if (link) {{ link.click(); }}
            </script>
            </body></html>
            """, height=0)
            st.session_state.downloaded_once = True

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            "⬇  Download Again",
            st.session_state.xlsx_bytes,
            file_name=st.session_state.xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption(
            "The Excel report downloads automatically when it's ready. "
            "If your browser blocked the auto-download, use the button above."
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
                <li>Upload A-direction SOR files (required) and B-direction (optional) as a ZIP or individual files</li>
                <li>Adjust the reburn threshold if needed - default is 0.150 dB</li>
                <li>Use the <strong>Include in Report</strong> checkboxes to filter which event types appear in the output</li>
                <li>Click <strong>Generate Report</strong> to download the color-coded Excel splice QC report</li>
            </ul>
        </div>
        <div class="tc-card" style="flex:1; margin-bottom:0;">
            <div class="tc-card-title">Excel Report Color Key</div>
            <div class="tc-legend">
                <span class="tc-pill"><span class="tc-swatch" style="background:#FFC7CE"></span>Pink - A+B Reburn</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FF4444"></span>Red - Break</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FF8800"></span>Orange - Broke</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#BDD7EE"></span>Blue - B-fill</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FFF2CC"></span>Yellow - A-only</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#FFD700"></span>Gold - A-only &#9888;</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#E8D5F5"></span>Lavender - B-only</span>
                <span class="tc-pill"><span class="tc-swatch" style="background:#C084FC"></span>Purple - B-only &#9888;</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
