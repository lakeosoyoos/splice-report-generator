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


def _detect_sites(dir_a: str, dir_b: str | None = None) -> tuple[str, str]:
    """Auto-detect short site codes from folder name or first SOR/JSON
    filename.  Falls back to ('A', 'B').  Best-effort — no required."""
    import re
    candidates: list[str] = []
    candidates.append(os.path.basename(os.path.normpath(dir_a)).upper())
    try:
        first = sorted(f for f in os.listdir(dir_a)
                       if f.lower().endswith((".sor", ".json")))
        if first:
            candidates.append(os.path.splitext(first[0])[0].upper())
    except Exception:
        pass
    for c in candidates:
        m = re.match(r"^([A-Z]{3,4})([A-Z]{3,4})", c)
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
with st.sidebar:
    st.header("Settings")

    ribbon_size = st.number_input(
        "Ribbon size (fibers per ribbon row)",
        min_value=1, max_value=24, value=int(engine.RIBBON_SIZE), step=1,
        help="Number of fibers per row in the output grid.  Standard cable "
             "ribbons are 12; some 432-count cables use 24 or other sizes.",
    )

    st.divider()
    with st.expander("Detection thresholds", expanded=True):
        reburn_thr = st.slider(
            "Reburn threshold — bidir loss to flag (dB)",
            min_value=0.050, max_value=0.500,
            value=float(engine.REBURN_THRESHOLD), step=0.005, format="%.3f",
            help="The bidirectional-average splice loss at or above which "
                 "a splice is flagged for reburn.  Default 0.160 dB.",
        )
        bend_thr = st.slider(
            "Bend threshold — minimum loss to call a bend (dB)",
            min_value=0.020, max_value=0.300,
            value=float(engine.BEND_THRESHOLD), step=0.005, format="%.3f",
            help="Positive loss step required for an off-splice event to "
                 "be classified as a bend.  Default 0.090 dB.",
        )
        closure_match_m = st.slider(
            "Splice match radius (m)",
            min_value=50, max_value=500,
            value=int(engine.CLOSURE_MATCH_KM * 1000), step=25,
            help="An event within this distance of a validated splice "
                 "center renders in the splice column; beyond it → bend / "
                 "damage column.  Default 150 m.",
        )
        position_tol_km = st.slider(
            "A↔B event-match tolerance (km)",
            min_value=0.3, max_value=3.0,
            value=float(engine.POSITION_TOL), step=0.1, format="%.1f",
            help="Maximum offset between an A-side event and its mirror in "
                 "the B-direction trace.  Default 1.5 km.",
        )

    with st.expander("End zones (launch / tailbox)", expanded=False):
        launch_fiber_max_km = st.slider(
            "Launch / tailbox zone size (km)",
            min_value=0.5, max_value=10.0,
            value=float(engine.LAUNCH_FIBER_MAX), step=0.5, format="%.1f",
            help="Events within this distance of the launch (km 0) or the "
                 "EOL are excluded from the splice / bend passes and "
                 "handled by the launch-issue detector instead.  Default "
                 "3.0 km.",
        )
        end_region_km = st.slider(
            "Cable-end region for phantom-closure drop (km)",
            min_value=0.5, max_value=10.0,
            value=float(engine.END_REGION_KM), step=0.5, format="%.1f",
            help="Phantom splice closures within this distance of the "
                 "estimated cable end are dropped.  Default 3.0 km.",
        )

    with st.expander("Launch / tailbox reflectance", expanded=False):
        bad_refl_db = st.slider(
            "Bad-reflectance threshold (dB)",
            min_value=-60.0, max_value=-30.0,
            value=float(engine.LAUNCH_BAD_REFL_DB), step=0.1, format="%.1f",
            help="A launch or tailbox connector whose reflectance is at or "
                 "above this value (closer to zero) flags as bad.  Default "
                 "−49.9 dB.  Healthy buried connectors reflect −50 to "
                 "−55 dB.",
        )
        tailbox_outlier_db = st.slider(
            "Tailbox population-outlier margin (dB)",
            min_value=2.0, max_value=25.0,
            value=10.0, step=1.0, format="%.0f",
            help="A fiber's tailbox refl must also be this many dB worse "
                 "than the per-direction population median to flag — "
                 "stops every fiber from lighting up when an entire span "
                 "was shot with the same bare-glass cable end.  Default "
                 "10 dB.",
        )

    with st.expander("Per-fiber bend gate", expanded=False):
        bend_res_splice_m = st.slider(
            "Splice residual (≤ this → splice) (m)",
            min_value=10, max_value=200,
            value=int(engine.BEND_RES_SPLICE_M), step=5,
            help="Per-fiber leave-one-out length-model residual at or below "
                 "this value classifies the event as the fiber's own "
                 "splice (just offset from the ribbon median).  Default "
                 "50 m.",
        )
        bend_res_bend_m = st.slider(
            "Bend residual (≥ this → bend candidate) (m)",
            min_value=50, max_value=500,
            value=int(engine.BEND_RES_BEND_M), step=10,
            help="Residual at or above this value qualifies the event for "
                 "bend classification (then confirmed by narrow-LSA at "
                 "predicted km).  Default 150 m.",
        )
        bend_narrow_loss_db = st.slider(
            "Narrow-LSA loss to confirm bend (dB)",
            min_value=0.005, max_value=0.100,
            value=float(engine.BEND_NARROW_LOSS_DB), step=0.005, format="%.3f",
            help="When residual lands in the bend zone, a narrow-LSA "
                 "reading at the predicted splice km must show at least "
                 "this much loss to confirm the bend is real.  Default "
                 "0.030 dB.",
        )

    st.divider()
    st.caption(
        "Slider values are applied per-run and reset to defaults on "
        "reload.  Defaults match the tech-reference report tuning."
    )


# Build the overrides dict in one place — what gets pushed onto the engine
# module before each run.
THRESHOLDS = {
    "REBURN_THRESHOLD":   float(reburn_thr),
    "BEND_THRESHOLD":     float(bend_thr),
    "CLOSURE_MATCH_KM":   float(closure_match_m) / 1000.0,
    "POSITION_TOL":       float(position_tol_km),
    "LAUNCH_FIBER_MAX":   float(launch_fiber_max_km),
    "END_REGION_KM":      float(end_region_km),
    "LAUNCH_BAD_REFL_DB": float(bad_refl_db),
    "BEND_RES_SPLICE_M":  int(bend_res_splice_m),
    "BEND_RES_BEND_M":    int(bend_res_bend_m),
    "BEND_NARROW_LOSS_DB": float(bend_narrow_loss_db),
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
