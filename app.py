"""
Splice Report Generator — Streamlit App
========================================
Bidirectional splice QC report from OTDR SOR or JSON files (EXFO FastReporter).

Simplified UI modelled after the Unidirectional One Shot app:
  • Sidebar: Settings (ribbon size + all detection thresholds as sliders)
  • Main:    1. Upload  →  2. Run  →  3. Results

Launch:  streamlit run app.py
"""
from __future__ import annotations

import os
import sys
import io
import tempfile
from contextlib import redirect_stdout

import numpy as np
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import splicereportmatchexfo as engine
from splicereportmatchexfo import (
    load_all, discover_splices, refine_closure_centers, detect_launch_issues,
    analyze_all, scan_a_standalone_events, scan_b_past_breaks,
    apply_field_gainer_rule,
    build_ribbon_data, write_xlsx,
    split_offsplice_events_into_own_columns,
    _normalize_untrimmed_events,
)


# ─────────────────────────────────────────────────────────────────────────────
#  Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Splice Report Generator",
    page_icon="🪢",
    layout="wide",
)

st.title("Splice Report Generator")
st.caption(
    "Bidirectional splice QC for OTDR SOR / JSON acquisitions.  Upload one "
    "ZIP (or any combination of files) per direction, click Run, get a "
    "ribbon-grid Excel report plus Zach's per-cell explanation PDF."
)


# ─────────────────────────────────────────────────────────────────────────────
#  Password gate (only enforced when st.secrets has one configured)
# ─────────────────────────────────────────────────────────────────────────────
def _check_password() -> bool:
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


if not _check_password():
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────
def _stage_uploads(uploaded_files, prefix: str) -> tuple[str, int, list[str]]:
    """Write uploaded files (.sor / .json / .zip) to a temp dir and return
    (dir_path, n_files_landed, [warnings]).  ZIPs are unpacked in place."""
    tmp = tempfile.mkdtemp(prefix=prefix)
    n = 0
    warns: list[str] = []
    import zipfile
    for uf in uploaded_files or []:
        name = uf.name
        lower = name.lower()
        try:
            uf.seek(0)
        except Exception:
            pass
        data = uf.read()
        if lower.endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
                    for nm in zf.namelist():
                        nm_l = nm.lower()
                        if not (nm_l.endswith(".sor") or nm_l.endswith(".json")):
                            continue
                        if nm.startswith("__MACOSX") or "/.DS_Store" in nm:
                            continue
                        base = os.path.basename(nm)
                        if not base:
                            continue
                        with zf.open(nm) as src, open(os.path.join(tmp, base), "wb") as dst:
                            dst.write(src.read())
                        n += 1
            except zipfile.BadZipFile as exc:
                warns.append(f"Could not read '{name}' as a ZIP ({exc}).")
        elif lower.endswith(".sor") or lower.endswith(".json"):
            with open(os.path.join(tmp, name), "wb") as fh:
                fh.write(data)
            n += 1
        else:
            warns.append(f"Skipping '{name}' (unsupported extension).")
    return tmp, n, warns


def _read_sor_genparams(filepath: str) -> dict:
    """Pull GenParams metadata (orig_loc, term_loc, ...) from a SOR
    file.  Ports the proven block-directory walk from
    unidirectional-one-shot/unidirectional_event_finder.py — the
    simple find('GenParams\\x00') approach doesn't get the right
    offset because GenParams block content lives AFTER the Map block
    based on declared sizes, not at the directory-entry position."""
    import struct
    try:
        with open(filepath, "rb") as f:
            data = f.read()

        # Walk the block directory using the Bellcore Map convention
        name_end = data.index(b"\x00", 0) + 1            # 'Map\x00'
        map_size = struct.unpack_from("<I", data, name_end + 2)[0]
        off = name_end + 6 + 2                            # skip num_blocks
        entries = []
        while off < map_size:
            ne = data.index(b"\x00", off) + 1
            nm = data[off:ne - 1].decode("latin-1")
            bv = struct.unpack_from("<H", data, ne)[0]
            bs = struct.unpack_from("<I", data, ne + 2)[0]
            entries.append((nm, bv, bs))
            off = ne + 6
        # Block content offsets accumulate from map_size
        cur = map_size
        body = None
        for nm, _bv, bs in entries:
            if nm == "GenParams":
                body = cur + len(nm) + 1
                break
            cur += bs
        if body is None:
            return {}

        # Walk the GenParams body field-by-field
        o = body + 2   # skip 2-byte language code

        def pull(o):
            e = data.index(b"\x00", o)
            return data[o:e].decode("latin-1", errors="replace").strip(), e + 1

        cable_id,   o = pull(o)
        fiber_id,   o = pull(o)
        o += 4   # FiberType (2) + NominalWavelength (2)
        orig_loc,   o = pull(o)
        term_loc,   o = pull(o)
        return {
            "cable_id": cable_id, "fiber_id": fiber_id,
            "orig_loc": orig_loc, "term_loc": term_loc,
        }
    except Exception:
        return {}


def _read_json_genparams(filepath: str) -> dict:
    """Same shape as _read_sor_genparams, but for EXFO JSON exports."""
    try:
        import json
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            obj = json.load(f)
        # JSON formats vary; check a few common nesting points
        for key_set in (obj, obj.get("GenParams") or {},
                         obj.get("identification") or {},
                         obj.get("Identification") or {}):
            o = (key_set.get("originating_location")
                 or key_set.get("orig_loc")
                 or key_set.get("OriginatingLocation"))
            t = (key_set.get("terminating_location")
                 or key_set.get("term_loc")
                 or key_set.get("TerminatingLocation"))
            if o or t:
                return {"orig_loc": (o or "").strip(),
                        "term_loc": (t or "").strip()}
        return {}
    except Exception:
        return {}


def _detect_sites(dir_a: str, dir_b: str | None = None) -> tuple[str, str]:
    """Best-effort detection of site_a / site_b names, in this order:
      1. orig_loc / term_loc from the first SOR/JSON file's GenParams
         (gives us full location names like 'MILLER' / 'TOPEKA').
      2. Filename SITE1+SITE2 pattern (e.g. MILTOP0001 → MIL + TOP).
      3. Folder-name SITE1+SITE2 pattern.
    Falls back to ('A','B') if nothing parses."""
    import re

    # ── 1.  Read GenParams metadata from the first file in dir_a ──
    try:
        files_a = sorted(f for f in os.listdir(dir_a)
                         if f.lower().endswith((".sor", ".json")))
    except Exception:
        files_a = []
    if files_a:
        path = os.path.join(dir_a, files_a[0])
        md = (_read_sor_genparams(path) if path.lower().endswith(".sor")
              else _read_json_genparams(path))
        orig = (md.get("orig_loc") or "").strip()
        term = (md.get("term_loc") or "").strip()
        if orig and term:
            # Strip stray quoting / whitespace.  Keep original casing
            # for display unless the value is all-numeric or empty.
            if not orig.isdigit() and not term.isdigit():
                return orig, term

    # ── 2 / 3.  Fall back to filename, then folder, regex matching ──
    filename_candidates: list[str] = []
    folder_candidates:   list[str] = [
        os.path.basename(os.path.normpath(dir_a)).upper()
    ]
    if files_a:
        filename_candidates.append(os.path.splitext(files_a[0])[0].upper())

    # Pattern requires the two site codes to be followed by a digit —
    # this is what every SOR filename in our corpus uses
    # (MILTOP0001, ELMMIL0541, SANDUR001, ...).
    strict = re.compile(r"^([A-Z]{3,4})([A-Z]{3,4})(?=\d)")
    for c in filename_candidates + folder_candidates:
        m = strict.match(c)
        if m:
            return m.group(1), m.group(2)
    # Lenient fallback (old behaviour) — only when nothing strict matched.
    lenient = re.compile(r"^([A-Z]{3,4})([A-Z]{3,4})")
    for c in filename_candidates + folder_candidates:
        m = lenient.match(c)
        if m:
            return m.group(1), m.group(2)
    return "A", "B"


def _apply_overrides(thresholds: dict) -> dict:
    """Snapshot the engine's module-level constants we're about to override,
    apply the slider values, and return the snapshot so the caller can
    restore them after the run."""
    saved = {k: getattr(engine, k) for k in thresholds}
    for k, v in thresholds.items():
        setattr(engine, k, v)
    return saved


def _restore_overrides(saved: dict) -> None:
    for k, v in saved.items():
        setattr(engine, k, v)


# ─────────────────────────────────────────────────────────────────────────────
#  Sidebar — Settings
# ─────────────────────────────────────────────────────────────────────────────
#  OTDR settings panel — matches the EXFO threshold-table layout
#  (Description / Apply / Fail / Warning columns).  Only the rows we
#  have engine code for actually do anything when their Apply checkbox
#  is ticked — see SUPPORTED below.
OTDR_ROWS = [
    # (key,                       label,                       fail_default,  unit,    supported)
    ("unidir_splice_loss",        "Unidir. splice loss",        0.250,        "dB",    True),
    ("bidir_splice_loss",         "Bidir splice loss",          0.160,        "dB",    True),
    ("unidir_connector_loss",     "Unidir. connector loss",     0.750,        "dB",    False),
    ("bidir_connector_loss",      "Bidir connector loss",       0.750,        "dB",    False),
    ("splitter_loss",             "Splitter Loss",              4.500,        "dB",    False),
    ("reflectance",               "Reflectance",                -49.9,        "dB",    True),
    ("fiber_section_atten",       "Fiber section attenuation",  0.400,        "dB/km", False),
    ("span_loss",                 "Span loss",                  20.000,       "dB",    False),
    ("span_length",               "Span length",                0.0000,       "km",    False),
    ("span_orl",                  "Span ORL",                   15.00,        "dB",    False),
]
# Pre-checked rows (match what splice report flags out of the box):
OTDR_DEFAULT_APPLY = {"unidir_splice_loss", "bidir_splice_loss", "reflectance"}

# Initialise persisted settings on first run
if "otdr_settings" not in st.session_state:
    st.session_state.otdr_settings = {
        key: {
            "apply":   key in OTDR_DEFAULT_APPLY,
            "fail":    fail,
            "warning": fail,        # Warning column captured visually; engine uses Fail only
        }
        for key, _, fail, _, _ in OTDR_ROWS
    }

with st.sidebar:
    # Every engine threshold the user can't reach via the OTDR panel
    # falls back to the module default — fixed for the duration of the run.
    ribbon_size         = int(engine.RIBBON_SIZE)
    bend_thr            = float(engine.BEND_THRESHOLD)
    closure_match_m     = int(engine.CLOSURE_MATCH_KM * 1000)
    position_tol_km     = float(engine.POSITION_TOL)
    min_pop_fraction    = int(round(engine.MIN_POP_FRACTION * 100))
    launch_fiber_max_km = float(engine.LAUNCH_FIBER_MAX)
    end_region_km       = float(engine.END_REGION_KM)
    spans_have_tailbox  = True
    tailbox_outlier_db  = 10.0
    bend_res_splice_m   = int(engine.BEND_RES_SPLICE_M)
    bend_res_bend_m     = int(engine.BEND_RES_BEND_M)
    bend_narrow_loss_db = float(engine.BEND_NARROW_LOSS_DB)

# ── OTDR settings panel — pixel-perfect EXFO match (Option B) ────────
# Rendered by a custom Streamlit component (raw HTML + CSS + JS) so we
# bypass Streamlit's widget chrome entirely.  Always visible in the
# sidebar.  See components/otdr_settings/index.html for the visual
# implementation.
from components.otdr_settings import otdr_settings as otdr_settings_component

with st.sidebar:
    # Widen the sidebar so the EXFO-styled table fits cleanly.
    st.markdown("""
    <style>
      section[data-testid="stSidebar"],
      section[data-testid="stSidebar"][aria-expanded="true"] {
        width: 620px !important;
        min-width: 620px !important;
        max-width: 620px !important;
      }
      section[data-testid="stSidebar"] > div {
        width: 620px !important;
        min-width: 620px !important;
      }
    </style>
    """, unsafe_allow_html=True)

    # Build the rows definition for the component.  Each row's initial
    # values come from session_state (the user's last-committed
    # settings); the supported flag tells the component whether to
    # decorate it as 'not yet wired'.
    _otdr_rows_for_component = [
        {
            "key":       key,
            "label":     label,
            "unit":      unit,
            "supported": supported,
            "initial":   st.session_state.otdr_settings[key],
        }
        for key, label, _fail, unit, supported in OTDR_ROWS
    ]
    _commit = otdr_settings_component(
        _otdr_rows_for_component,
        default=None,
        key="otdr_component",
    )
    if _commit:
        # User clicked Apply settings inside the component — persist
        # to session_state.otdr_settings for the next report run.
        for key, vals in _commit.items():
            st.session_state.otdr_settings[key] = {
                "apply":   bool(vals.get("apply")),
                "fail":    float(vals.get("fail", 0.0)),
                "warning": float(vals.get("warning", 0.0)),
            }


# OTDR settings overrides — when a row's Apply checkbox is ticked,
# the Fail value overrides the engine default for that threshold.
# Unticked rows fall back to the engine value.
otdr = st.session_state.get("otdr_settings", {})
def _otdr_override(key, engine_default):
    row = otdr.get(key) or {}
    if row.get("apply") and row.get("fail") is not None:
        return float(row["fail"])
    return engine_default

reburn_thr     = _otdr_override("bidir_splice_loss",  float(engine.REBURN_THRESHOLD))
single_dir_thr = _otdr_override("unidir_splice_loss", float(engine.SINGLE_DIR_THRESHOLD))
bad_refl_db    = _otdr_override("reflectance",        float(engine.LAUNCH_BAD_REFL_DB))

# Build the overrides dict in one place — what gets pushed onto the engine
# module before each run.
THRESHOLDS = {
    "REBURN_THRESHOLD":     float(reburn_thr),
    "SINGLE_DIR_THRESHOLD": float(single_dir_thr),
    "BEND_THRESHOLD":       float(bend_thr),
    "CLOSURE_MATCH_KM":     float(closure_match_m) / 1000.0,
    "POSITION_TOL":         float(position_tol_km),
    "MIN_POP_FRACTION":     float(min_pop_fraction) / 100.0,
    "LAUNCH_FIBER_MAX":     float(launch_fiber_max_km),
    "END_REGION_KM":        float(end_region_km),
    "LAUNCH_BAD_REFL_DB":   float(bad_refl_db),
    "BEND_RES_SPLICE_M":    int(bend_res_splice_m),
    "BEND_RES_BEND_M":      int(bend_res_bend_m),
    "BEND_NARROW_LOSS_DB":  float(bend_narrow_loss_db),
}
# TAILBOX_OUTLIER_DB is a local in detect_launch_issues — it's not a module
# constant — so we pass it through via a dedicated kwarg below.  Stash it
# here for reference.
TAILBOX_OUTLIER_DB = float(tailbox_outlier_db)


# ─────────────────────────────────────────────────────────────────────────────
#  Step 1 — Upload
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("1. Upload")
cols = st.columns(2)
with cols[0]:
    uploaded_a = st.file_uploader(
        "A-direction files",
        type=["sor", "json", "zip"],
        accept_multiple_files=True,
        help="Drop a ZIP of one direction's SOR / JSON files, or the loose "
             "files themselves.",
    )
    if uploaded_a:
        total_kb = sum(f.size for f in uploaded_a) / 1024
        st.caption(f"A: {len(uploaded_a)} file(s), {total_kb:.0f} KB total")
with cols[1]:
    uploaded_b = st.file_uploader(
        "B-direction files",
        type=["sor", "json", "zip"],
        accept_multiple_files=True,
        help="The opposite-direction shoot of the same span.  Required — "
             "single-direction analysis can't compute a bidirectional "
             "average.",
    )
    if uploaded_b:
        total_kb = sum(f.size for f in uploaded_b) / 1024
        st.caption(f"B: {len(uploaded_b)} file(s), {total_kb:.0f} KB total")

if not (uploaded_a and uploaded_b):
    st.info(
        "Drop in one folder zip (or the loose .sor / .json files) for each "
        "direction.  Direction is inferred from file metadata — naming is "
        "not enforced."
    )
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
#  Step 2 — Run
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("2. Run")
run_clicked = st.button("Generate Report", type="primary",
                        use_container_width=True)

# Reset previous-result flags when a new run is kicked off
if run_clicked:
    for key in ("xlsx_bytes", "xlsx_name", "zach_pdf_bytes", "zach_pdf_name",
                "summary_data", "done"):
        st.session_state[key] = None


# ─────────────────────────────────────────────────────────────────────────────
#  Engine run (only when the button is clicked this turn)
# ─────────────────────────────────────────────────────────────────────────────
if run_clicked:
    prog = st.progress(0.0, text="Staging A-direction files...")
    dir_a, n_a, warns_a = _stage_uploads(uploaded_a, "splice_a_")
    for w in warns_a:
        st.warning(w)
    if n_a == 0:
        prog.empty()
        st.error("No .sor or .json files found in the A-direction upload.")
        st.stop()

    prog.progress(0.10, text=f"Staging B-direction files ({n_a} A files staged)...")
    dir_b, n_b, warns_b = _stage_uploads(uploaded_b, "splice_b_")
    for w in warns_b:
        st.warning(w)
    if n_b == 0:
        prog.empty()
        st.error("No .sor or .json files found in the B-direction upload.")
        st.stop()
    prog.progress(0.18, text=f"Staged {n_a} A + {n_b} B files.")

    site_a, site_b = _detect_sites(dir_a, dir_b)
    log_buf = io.StringIO()

    # Apply slider overrides to the engine module for the duration of this run
    saved = _apply_overrides(THRESHOLDS)
    try:
        prog.progress(0.25, text="Loading OTDR files...")
        with redirect_stdout(log_buf):
            fibers_a, fibers_b = load_all(dir_a, dir_b)
        if not fibers_a:
            st.error("Loaded zero fibers from the A-direction upload.")
            prog.empty()
            st.stop()
        n_fibers = max(fibers_a.keys())

        # Pass-0: normalize untrimmed SOR shoots.  For SORs where the tech
        # didn't set start/stop, events[0] is the OTDR port at km 0 and
        # events[1] is the launch connector at km ~1.0.  Without this
        # normalize, the launch event leaks through as a "splice candidate"
        # at km 1 and gets reported as a bend column on every fiber.  Save
        # the original events on r['_raw_events'] so detect_launch_issues
        # can still find the actual tailbox 1F event (normalize strips it).
        prog.progress(0.30, text="Normalizing untrimmed SOR events...")
        with redirect_stdout(log_buf):
            for r in list(fibers_a.values()) + list(fibers_b.values()):
                r["_raw_events"] = r["events"]
                r["events"] = _normalize_untrimmed_events(r["events"])

        prog.progress(0.35, text="Discovering splice closures...")
        with redirect_stdout(log_buf):
            cand = discover_splices(fibers_a)
            real_sp, phantoms = refine_closure_centers(
                fibers_a, cand, return_phantoms=True)
            splices = sorted(
                list(real_sp) + list(phantoms),
                key=lambda sp: sp.get("position_km_refined", sp["position_km"]),
            )
            sdn = 0
            for sp in splices:
                if sp.get("column_kind") == "splice":
                    sdn += 1
                    sp["splice_display_num"] = sdn

        prog.progress(0.45, text="Detecting launch / tailbox issues...")
        first_splice_km = splices[0]["position_km"] if splices else None
        with redirect_stdout(log_buf):
            launch_issues = detect_launch_issues(
                fibers_a, fibers_b, first_splice_km,
                bad_refl_db=THRESHOLDS["LAUNCH_BAD_REFL_DB"],
                spans_have_tailbox=bool(spans_have_tailbox),
            )

        # Span estimate (median of top-quartile EOFs across A fibers)
        ends = sorted(e["dist_km"] for r in fibers_a.values()
                      for e in r["events"] if e.get("is_end"))
        actual_span = (round(float(np.median(ends[int(len(ends) * 0.75):])), 2)
                       if ends else 0)

        prog.progress(0.55, text="Pass 1 — at-splice classification...")
        with redirect_stdout(log_buf):
            results = analyze_all(fibers_a, fibers_b, splices,
                                  THRESHOLDS["REBURN_THRESHOLD"])

        prog.progress(0.65, text="Pass 2a — standalone A events (bends / breaks)...")
        with redirect_stdout(log_buf):
            a_standalone = scan_a_standalone_events(
                fibers_a, splices, results, actual_span, fibers_b=fibers_b,
            )

        # Pass 2b — bidirectional ghost reflections (mid-span 1F, near-zero
        # loss in BOTH directions).  Older builds of the engine don't have
        # this function; guard with hasattr so the app still runs.
        ghost_refl = {}
        if hasattr(engine, "scan_bidir_ghost_reflections"):
            prog.progress(0.72, text="Pass 2b — bidirectional ghost reflections...")
            with redirect_stdout(log_buf):
                ghost_refl = engine.scan_bidir_ghost_reflections(
                    fibers_a, fibers_b, splices,
                    {**results, **a_standalone}, actual_span,
                )

        # Pass 2b' — EXFO merged-reflective events (0F with negative refl)
        merged_refl = {}
        if hasattr(engine, "scan_merged_reflective_events"):
            prog.progress(0.78, text="Pass 2b' — EXFO merged-reflective events...")
            with redirect_stdout(log_buf):
                merged_refl = engine.scan_merged_reflective_events(
                    fibers_a, fibers_b, splices,
                    {**results, **a_standalone, **ghost_refl}, actual_span,
                )

        prog.progress(0.82, text="Pass 2c — past-break B-fill...")
        with redirect_stdout(log_buf):
            b_pastbreak = scan_b_past_breaks(
                fibers_a, fibers_b, splices, THRESHOLDS["REBURN_THRESHOLD"],
                results, actual_span,
            )

        all_results = {**results, **a_standalone, **ghost_refl,
                       **merged_refl, **b_pastbreak}

        with redirect_stdout(log_buf):
            apply_field_gainer_rule(all_results, actual_span)
            # Off-splice events into their own bend/damage columns
            all_results, splices = split_offsplice_events_into_own_columns(
                all_results, splices, total_span_km=actual_span,
            )

        prog.progress(0.90, text="Building ribbon grid + writing Excel...")
        with redirect_stdout(log_buf):
            cells, launch_cells_a, launch_cells_b = build_ribbon_data(
                all_results, n_fibers, int(ribbon_size), len(splices),
                launch_issues=launch_issues,
            )

        xlsx_dir = tempfile.mkdtemp(prefix="splice_xlsx_")
        xlsx_path = os.path.join(xlsx_dir, "splice_report.xlsx")
        with redirect_stdout(log_buf):
            write_xlsx(cells, splices, n_fibers, int(ribbon_size), xlsx_path,
                       site_a, site_b, actual_span,
                       launch_cells_a=launch_cells_a,
                       launch_cells_b=launch_cells_b)
        with open(xlsx_path, "rb") as fh:
            xlsx_bytes = fh.read()

        # Zach's explanation PDF (best effort — don't block xlsx delivery)
        zach_bytes = None
        zach_name = None
        try:
            from make_zach_explanation import build_explanation_pdf
            zach_path = os.path.join(xlsx_dir, "zach_explanation.pdf")
            with redirect_stdout(log_buf):
                build_explanation_pdf(
                    all_results, splices, launch_issues, actual_span,
                    site_a, site_b, zach_path,
                    ribbon_size=int(ribbon_size),
                    reburn_threshold=THRESHOLDS["REBURN_THRESHOLD"],
                )
            with open(zach_path, "rb") as fh:
                zach_bytes = fh.read()
            zach_name = f"zach_explanation_{site_a}_{site_b}.pdf"
        except Exception as exc:
            log_buf.write(f"\n[zach-explanation] failed: {exc}\n")

        # Summary numbers for the metrics row
        n_reburn = sum(1 for r in all_results.values()
                       if r.get("event_source") in ("bidir", "bidir_grey_a", "bidir_grey_b")
                       and not r.get("is_break") and not r.get("is_bend"))
        n_break  = sum(1 for r in all_results.values() if r.get("is_break"))
        n_broke  = sum(1 for r in all_results.values() if r.get("is_broke"))
        n_bend   = sum(1 for r in all_results.values() if r.get("is_bend"))
        n_ref    = sum(1 for r in all_results.values() if r.get("is_ref"))
        n_launch = len(launch_issues)
        n_total  = len(all_results)

        st.session_state.xlsx_bytes = xlsx_bytes
        st.session_state.xlsx_name = f"splice_report_{site_a}_{site_b}.xlsx"
        st.session_state.zach_pdf_bytes = zach_bytes
        st.session_state.zach_pdf_name = zach_name
        st.session_state.summary_data = dict(
            site_a=site_a, site_b=site_b,
            n_fibers=n_fibers, n_splices=len(splices), span=actual_span,
            n_reburn=n_reburn, n_break=n_break, n_broke=n_broke,
            n_bend=n_bend, n_ref=n_ref, n_launch=n_launch, n_total=n_total,
        )
        st.session_state.done = True
        prog.progress(1.0, text="Done.")
        prog.empty()
    finally:
        _restore_overrides(saved)


# ─────────────────────────────────────────────────────────────────────────────
#  Step 3 — Results
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.get("done"):
    d = st.session_state.summary_data
    st.subheader(f"3. Results — {d['site_a']} → {d['site_b']}")

    m = st.columns(7)
    m[0].metric("Fibers",       d["n_fibers"])
    m[1].metric("Span (km)",    f"{d['span']:.2f}")
    m[2].metric("Splices",      d["n_splices"])
    m[3].metric("Reburns",      d["n_reburn"])
    m[4].metric("Bends",        d["n_bend"])
    m[5].metric("Broke / Break", d["n_broke"] + d["n_break"])
    m[6].metric("Launch issues", d["n_launch"])

    st.caption(f"Total flagged events: **{d['n_total']}**  |  "
               f"Ref events: {d['n_ref']}")

    st.divider()

    dl_cols = st.columns(2)
    with dl_cols[0]:
        st.download_button(
            "⬇  Download Excel report",
            data=st.session_state.xlsx_bytes,
            file_name=st.session_state.xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with dl_cols[1]:
        if st.session_state.zach_pdf_bytes:
            st.download_button(
                "⬇  Download Zach's Explanation PDF",
                data=st.session_state.zach_pdf_bytes,
                file_name=st.session_state.zach_pdf_name,
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.caption("Zach's explanation PDF unavailable for this run "
                       "(see log).")
else:
    st.info("Click **Generate Report** to process the uploaded files.")
