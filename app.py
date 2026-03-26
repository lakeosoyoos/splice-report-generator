"""
Splice Report Generator — Streamlit App
========================================
Generate splice QC reports from OTDR SOR files (bidirectional only, no uni).

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

from splice_report_generator import (
    load_all, discover_splices, analyze_all, build_ribbon_data, write_xlsx,
    REBURN_THRESHOLD, NOMINAL_SPLICE, RIBBON_SIZE,
)

# ── Page config ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="ZERO dB — Splice Report Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stButton > button[kind="primary"],
    .stDownloadButton > button[kind="primary"] {
        background-color: #4BA82E !important;
        border-color: #4BA82E !important;
        color: white !important;
    }
    .stButton > button[kind="primary"]:hover,
    .stDownloadButton > button[kind="primary"]:hover {
        background-color: #3D8C24 !important;
        border-color: #3D8C24 !important;
    }
    .stButton > button,
    .stDownloadButton > button {
        border-color: #4BA82E !important;
        color: #4BA82E !important;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: #3D8C24 !important;
        color: #3D8C24 !important;
    }
    .stProgress > div > div > div > div {
        background-color: #4BA82E !important;
    }
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: white !important;
    }
    .stRadio [role="radiogroup"] label[data-checked="true"],
    .stRadio [role="radiogroup"] label:has(input:checked) {
        background-color: #4BA82E !important;
        border-color: #4BA82E !important;
        color: white !important;
    }
    .stRadio [role="radiogroup"] label[data-checked="true"] p,
    .stRadio [role="radiogroup"] label:has(input:checked) p {
        color: white !important;
    }
    .stRadio [role="radiogroup"] label {
        border-color: #4BA82E !important;
    }
    a { color: #4BA82E !important; }
</style>
""", unsafe_allow_html=True)

st.title("Splice Report Generator")
st.caption("Bidirectional splice QC report from OTDR SOR files (no unidirectional)")


# ── Password protection ──────────────────────────────────────────────────────

def check_password():
    """Return True if user entered the correct password."""
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


# ── Session state ────────────────────────────────────────────────────────────

for key in ["xlsx_bytes", "xlsx_name", "summary", "log_output", "done"]:
    if key not in st.session_state:
        st.session_state[key] = None
if "upload_key" not in st.session_state:
    st.session_state.upload_key = 0


# ── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("Upload SOR Files")

    input_method = st.radio(
        "Input method",
        ["Upload ZIP", "Browse files", "Folder path"],
        index=0,
        horizontal=True,
    )

    uploaded_a = None
    uploaded_b = None
    zip_a = None
    zip_b = None
    folder_a = None
    folder_b = None

    if input_method == "Upload ZIP":
        zip_a = st.file_uploader(
            "A-direction ZIP",
            type=["zip"],
            accept_multiple_files=False,
            key=f"zip_a_{st.session_state.upload_key}",
        )
        if zip_a:
            st.caption(f"A: {zip_a.name} ({zip_a.size / 1024:.0f} KB)")
        zip_b = st.file_uploader(
            "B-direction ZIP (optional)",
            type=["zip"],
            accept_multiple_files=False,
            key=f"zip_b_{st.session_state.upload_key}",
        )
        if zip_b:
            st.caption(f"B: {zip_b.name} ({zip_b.size / 1024:.0f} KB)")
    elif input_method == "Browse files":
        uploaded_a = st.file_uploader(
            "A-direction SOR files",
            type=["sor"],
            accept_multiple_files=True,
            key=f"upload_a_{st.session_state.upload_key}",
        )
        uploaded_b = st.file_uploader(
            "B-direction SOR files (optional)",
            type=["sor"],
            accept_multiple_files=True,
            key=f"upload_b_{st.session_state.upload_key}",
        )
    else:
        folder_a = st.text_input(
            "A-direction folder path",
            value=st.session_state.get("folder_a", ""),
            placeholder="/Users/you/Desktop/A Direction/",
        )
        if folder_a:
            folder_a = folder_a.strip().strip("'\"")
            st.session_state.folder_a = folder_a
            if os.path.isdir(folder_a):
                n = len([f for f in os.listdir(folder_a) if f.lower().endswith('.sor')])
                st.caption(f"Found {n} .sor files")
            else:
                st.warning("Folder not found")

        folder_b = st.text_input(
            "B-direction folder path (optional)",
            value=st.session_state.get("folder_b", ""),
            placeholder="/Users/you/Desktop/B Direction/",
        )
        if folder_b:
            folder_b = folder_b.strip().strip("'\"")
            st.session_state.folder_b = folder_b
            if os.path.isdir(folder_b):
                n = len([f for f in os.listdir(folder_b) if f.lower().endswith('.sor')])
                st.caption(f"Found {n} .sor files")
            elif folder_b.strip():
                st.warning("Folder not found")

    if st.button("Clear All", use_container_width=True):
        old_key = st.session_state.upload_key
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.session_state.upload_key = old_key + 1
        st.rerun()

    st.divider()
    st.subheader("Settings")

    site_a = st.text_input("Site A name", value="STR")
    site_b = st.text_input("Site B name", value="ROM")
    threshold = st.number_input("Reburn threshold (dB)", value=REBURN_THRESHOLD,
                                format="%.3f", step=0.01)
    ribbon_size = st.number_input("Fibers per ribbon", value=RIBBON_SIZE,
                                   min_value=1, max_value=24, step=1)
    span_km = st.number_input("Span distance (km, 0=auto)", value=0.0,
                               format="%.2f", step=1.0)

    has_a = (bool(uploaded_a) or bool(zip_a) or
             (folder_a and os.path.isdir(folder_a)))
    run_button = st.button("Generate Report", type="primary",
                           use_container_width=True, disabled=not has_a)


# ── Helpers ──────────────────────────────────────────────────────────────────

def stage_files(uploaded, prefix="sor_"):
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    for uf in uploaded:
        fp = os.path.join(tmpdir, uf.name)
        with open(fp, 'wb') as f:
            f.write(uf.getbuffer())
    return tmpdir


def stage_zip(uploaded_zip, prefix="sor_zip_"):
    """Extract SOR files from a ZIP to a temp directory. Return directory path."""
    import zipfile
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    with zipfile.ZipFile(io.BytesIO(uploaded_zip.getbuffer()), 'r') as zf:
        for name in zf.namelist():
            if name.lower().endswith('.sor') and not name.startswith('__MACOSX'):
                basename = os.path.basename(name)
                if not basename:
                    continue
                fp = os.path.join(tmpdir, basename)
                with zf.open(name) as src, open(fp, 'wb') as dst:
                    dst.write(src.read())
    return tmpdir


# ── Run ──────────────────────────────────────────────────────────────────────

if run_button and has_a:
    # Get directories
    if folder_a and os.path.isdir(folder_a):
        dir_a = folder_a
        dir_b = folder_b if (folder_b and os.path.isdir(folder_b)) else None
    elif zip_a:
        progress = st.progress(0, text="Extracting A-direction ZIP...")
        dir_a = stage_zip(zip_a, "splice_a_")
        progress.progress(40, text="Extracting B-direction ZIP...")
        dir_b = stage_zip(zip_b, "splice_b_") if zip_b else None
        progress.progress(50, text="Files extracted.")
        progress.empty()
    else:
        progress = st.progress(0, text="Staging A-direction files...")
        dir_a = stage_files(uploaded_a, "splice_a_")
        progress.progress(40, text="Staging B-direction files...")
        dir_b = stage_files(uploaded_b, "splice_b_") if uploaded_b else None
        progress.progress(50, text="Files staged.")
        progress.empty()

    analysis_bar = st.progress(0, text="Loading SOR files...")

    log_buf = io.StringIO()
    with redirect_stdout(log_buf):
        fibers_a, fibers_b = load_all(dir_a, dir_b)

    n_fibers = max(fibers_a.keys()) if fibers_a else 0
    analysis_bar.progress(20, text=f"Loaded {len(fibers_a)} A + {len(fibers_b)} B fibers...")

    with redirect_stdout(log_buf):
        splices = discover_splices(fibers_a)
    analysis_bar.progress(40, text=f"Found {len(splices)} splice closures...")

    # Auto-detect span
    actual_span = span_km
    if actual_span == 0:
        ends = [e['dist_km'] for r in fibers_a.values()
                for e in r['events'] if e['is_end'] and e['dist_km'] > 90]
        actual_span = round(np.median(ends), 2) if ends else 97.33

    analysis_bar.progress(50, text=f"Analyzing {n_fibers} fibers at {len(splices)} splices...")
    with redirect_stdout(log_buf):
        results = analyze_all(fibers_a, fibers_b, splices, threshold)

    n_flagged = len(results)
    n_breaks = sum(1 for r in results.values() if r['is_break'])
    n_broke = sum(1 for r in results.values() if r['is_broke'])
    n_reburn = n_flagged - n_breaks - n_broke

    analysis_bar.progress(70, text="Building ribbon grid...")
    with redirect_stdout(log_buf):
        cells = build_ribbon_data(results, n_fibers, ribbon_size, len(splices))

    analysis_bar.progress(85, text="Writing Excel report...")
    xlsx_tmpdir = tempfile.mkdtemp(prefix="splice_xlsx_")
    xlsx_path = os.path.join(xlsx_tmpdir, "splice_report.xlsx")
    with redirect_stdout(log_buf):
        write_xlsx(cells, splices, n_fibers, ribbon_size, xlsx_path,
                   site_a, site_b, actual_span)

    with open(xlsx_path, 'rb') as f:
        st.session_state.xlsx_bytes = f.read()
    st.session_state.xlsx_name = f"splice_report_{site_a}_{site_b}.xlsx"

    summary = [
        f"**Fibers:** {n_fibers}",
        f"**Splice closures:** {len(splices)}",
        f"**Span:** {actual_span} km",
        f"**Threshold:** {threshold:.3f} dB",
        "",
        f"**Flagged events:** {n_flagged}",
        f"  - Breaks: {n_breaks}",
        f"  - Broke: {n_broke}",
        f"  - Reburns: {n_reburn}",
    ]
    st.session_state.summary = "\n\n".join(summary)
    st.session_state.log_output = log_buf.getvalue()
    st.session_state.done = True

    analysis_bar.progress(100, text="Done!")
    analysis_bar.empty()


# ── Display ──────────────────────────────────────────────────────────────────

if st.session_state.get("done"):
    st.subheader("Report Complete")
    if st.session_state.summary:
        st.markdown(st.session_state.summary)

    st.divider()

    if st.session_state.xlsx_bytes:
        st.download_button(
            "⬇ Download Excel Report",
            st.session_state.xlsx_bytes,
            file_name=st.session_state.xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    with st.expander("Analysis Log"):
        st.code(st.session_state.log_output or "No log.", language=None)

else:
    st.info("Upload A-direction SOR files (and optionally B-direction) in the sidebar, then click **Generate Report**.")
    st.markdown("""
    **How it works:**
    1. Upload **A-direction** SOR files (required) and **B-direction** (optional)
    2. Set your site names, reburn threshold, and ribbon size
    3. Click **Generate Report**
    4. Download the Excel splice QC report

    **Report contents:**
    - One row per 12-fiber ribbon, one column per splice closure
    - Bidirectional splice loss for flagged fibers
    - Breaks (Fresnel reflection), broke fibers, reburn candidates
    - No unidirectional entries
    """)
