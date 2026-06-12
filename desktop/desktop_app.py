"""
Splice Report Generator — Desktop UI
====================================
Folder-based Streamlit UI for the bundled Windows desktop app.  Mirrors
the web app's customer-profile dropdown + EXFO-style OTDR settings panel,
but reads two LOCAL folders (A-direction, B-direction) via a native
tkinter picker instead of upload widgets — no file-size limit, nothing
ever leaves the machine.

Boot order:
    launcher.exe → boots Streamlit headless → opens browser tab → this
    script runs.  The launcher pre-seeds STREAMLIT_SERVER_HEADLESS=true
    and binds to 127.0.0.1:8501 only.
"""
from __future__ import annotations

import io
import os
import sys
import tempfile
import shutil
import zipfile
from contextlib import redirect_stdout
from pathlib import Path

import streamlit as st


# ─────────────────────────────────────────────────────────────────────────────
#  Engine import — auto-update aware
# ─────────────────────────────────────────────────────────────────────────────
# The launcher sets SS_ENGINE_DIR when it has successfully downloaded a
# fresh engine into ~/.spliceReport/engine and validated it.  If set, we
# prepend that directory to sys.path so `import splicereportmatchexfo`
# resolves to the freshly-downloaded copy.  Otherwise we fall back to the
# bundled copy living next to the frozen exe (or alongside this file in
# dev mode).
_ENGINE_DIR = os.environ.get("SS_ENGINE_DIR")
if _ENGINE_DIR and os.path.isdir(_ENGINE_DIR):
    sys.path.insert(0, _ENGINE_DIR)
# Always also add the script's own directory so the bundled copy works
# during dev (running `streamlit run desktop/desktop_app.py`).
_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_HERE)
for p in (_HERE, _REPO_ROOT):
    if p not in sys.path:
        sys.path.insert(0, p)

import splicereportmatchexfo as engine  # noqa: E402
from splicereportmatchexfo import (  # noqa: E402
    load_all, discover_splices, refine_closure_centers, detect_launch_issues,
    analyze_all, scan_a_standalone_events, scan_b_past_breaks,
    scan_b_side_breaks, apply_field_gainer_rule, apply_connector_loss_rule,
    build_ribbon_data, write_xlsx,
    split_offsplice_events_into_own_columns,
    _normalize_untrimmed_events,
)


# ─────────────────────────────────────────────────────────────────────────────
#  Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Splice Report — Desktop",
    page_icon="🪢",
    layout="wide",
)
st.title("Splice Report Generator (Desktop)")
st.caption(
    "Bidirectional OTDR splice QC, running 100% on this machine.  "
    "Pick the A-direction and B-direction folders, click Run, and the "
    "report drops next to the inputs."
)


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────
SOR_MAGIC = b"Map\x00"             # Bellcore SR-4731 SOR signature
JSON_KEYS = (b'"events"', b'"genericParameters"', b'"GenParams"',
             b'"keyEvents"', b'"key_events"')


def _looks_like_sor(path: str) -> bool:
    """Quick byte sniff — first 8 bytes must contain the SR-4731 Map block
    header.  Real SOR files start with 'Map\\x00' optionally preceded by a
    short version prefix.  This rejects stray .sor-named results files
    that accidentally share the extension."""
    try:
        with open(path, "rb") as fh:
            head = fh.read(8)
        return SOR_MAGIC in head
    except OSError:
        return False


def _looks_like_json_otdr(path: str) -> bool:
    """Quick content sniff — read the first 4 KB and look for one of the
    JSON keys EXFO uses for an event-table export.  Rejects random JSON
    config files."""
    try:
        with open(path, "rb") as fh:
            head = fh.read(4096)
        return any(k in head for k in JSON_KEYS)
    except OSError:
        return False


def _is_otdr_file(path: str) -> bool:
    low = path.lower()
    if low.endswith(".sor"):
        return _looks_like_sor(path)
    if low.endswith(".json"):
        return _looks_like_json_otdr(path)
    return False


def _walk_otdr(folder: str) -> list[str]:
    out = []
    for root, _, files in os.walk(folder):
        for f in files:
            full = os.path.join(root, f)
            if _is_otdr_file(full):
                out.append(full)
    return out


def _stage_flat(paths: list[str], prefix: str) -> tuple[str, list[str]]:
    """Copy every path into a fresh temp dir with a flat layout.
    De-duplicates basenames by suffixing  __<n>  so two files with the same
    name living in different subfolders don't overwrite each other.
    Returns (temp_dir, [warnings])."""
    tmp = tempfile.mkdtemp(prefix=prefix)
    warns: list[str] = []
    seen: dict[str, int] = {}
    for src in paths:
        base = os.path.basename(src)
        stem, ext = os.path.splitext(base)
        n = seen.get(base, 0)
        if n == 0:
            dest_name = base
        else:
            dest_name = f"{stem}__{n}{ext}"
            warns.append(
                f"Duplicate basename '{base}' — staged as '{dest_name}'.")
        seen[base] = n + 1
        try:
            shutil.copy2(src, os.path.join(tmp, dest_name))
        except OSError as exc:
            warns.append(f"Skipped '{src}' ({exc}).")
    return tmp, warns


def _pick_folder(button_label: str, key: str) -> str | None:
    """Open the native folder picker and return the chosen path (or None).
    Uses tkinter — no extra deps.  Falls back to the paste-path box if
    tkinter isn't available (very rare, but possible on some Linux
    runners)."""
    if st.button(button_label, key=key):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            path = filedialog.askdirectory(parent=root)
            root.destroy()
            return path or None
        except Exception as exc:
            st.warning(f"Native folder picker unavailable ({exc}). "
                       f"Paste the path below instead.")
            return None
    return None


# ─────────────────────────────────────────────────────────────────────────────
#  Customer profiles + OTDR settings (mirror of the web app)
# ─────────────────────────────────────────────────────────────────────────────
OTDR_ROWS = [
    ("unidir_splice_loss",        "Unidir. splice loss",        0.250,        "dB",    True),
    ("bidir_splice_loss",         "Bidir splice loss",          0.160,        "dB",    True),
    ("unidir_connector_loss",     "Unidir. connector loss",     0.750,        "dB",    False),
    ("bidir_connector_loss",      "Bidir connector loss",       0.500,        "dB",    True),
    ("splitter_loss",             "Splitter Loss",              4.500,        "dB",    False),
    ("reflectance",               "Reflectance",                -49.9,        "dB",    True),
    ("fiber_section_atten",       "Fiber section attenuation",  0.400,        "dB/km", False),
    ("span_loss",                 "Span loss",                  20.000,       "dB",    False),
    ("span_length",               "Span length",                0.0000,       "km",    False),
    ("span_orl",                  "Span ORL",                   15.00,        "dB",    False),
]
OTDR_DEFAULT_APPLY = {"unidir_splice_loss", "bidir_splice_loss",
                      "bidir_connector_loss", "reflectance"}

CUSTOMER_PROFILES = {
    "Default (engine baseline)": {"apply": set(OTDR_DEFAULT_APPLY),
                                   "thresholds": {}},
    "Lumen": {
        "apply": {"unidir_splice_loss", "bidir_splice_loss",
                   "bidir_connector_loss", "reflectance"},
        "thresholds": {"bidir_splice_loss": 0.120,
                       "unidir_splice_loss": 0.200,
                       "bidir_connector_loss": 0.400,
                       "reflectance": -50.0},
    },
    "Zayo": {
        "apply": {"bidir_splice_loss", "bidir_connector_loss"},
        "thresholds": {"bidir_splice_loss": 0.200,
                       "bidir_connector_loss": 0.600},
    },
    "Custom (edit table below)": {"apply": None, "thresholds": None},
}


def _settings_from_profile(name: str) -> dict:
    prof = CUSTOMER_PROFILES.get(name) or {}
    apply_set = prof.get("apply")
    overrides = prof.get("thresholds") or {}
    out = {}
    for key, _, fail_default, _, _ in OTDR_ROWS:
        fail = float(overrides.get(key, fail_default))
        applied = ((apply_set is not None and key in apply_set)
                   if apply_set is not None
                   else (key in OTDR_DEFAULT_APPLY))
        out[key] = {"apply": applied, "fail": fail, "warning": fail}
    return out


_profile_names = list(CUSTOMER_PROFILES.keys())
if st.session_state.get("otdr_profile") not in _profile_names:
    st.session_state.otdr_profile = _profile_names[0]
if "otdr_settings" not in st.session_state:
    st.session_state.otdr_settings = _settings_from_profile(
        st.session_state.otdr_profile)


# ─────────────────────────────────────────────────────────────────────────────
#  Sidebar
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Splice Report — Desktop")

    src = os.environ.get("SS_ENGINE_SOURCE", "bundled (offline)")
    st.caption(f"Engine: **{src}**")

    if st.button("Quit app", type="secondary",
                 help="Stops the local Streamlit server and closes."):
        # Hard-exit so PyInstaller's launcher subprocess dies too.
        os._exit(0)

    st.markdown("---")
    st.markdown("**Customer profile**")
    _cur = st.session_state.otdr_profile
    _picked = st.selectbox(
        "Customer",
        _profile_names,
        index=_profile_names.index(_cur),
        label_visibility="collapsed",
        key="otdr_profile_select_desktop",
    )
    if _picked != _cur:
        st.session_state.otdr_profile = _picked
        if "Custom" not in _picked:
            st.session_state.otdr_settings = _settings_from_profile(_picked)
        st.rerun()

    st.markdown("---")
    st.markdown("**OTDR thresholds**")
    for key, label, default, unit, supported in OTDR_ROWS:
        row = st.session_state.otdr_settings[key]
        c1, c2 = st.columns([3, 2])
        with c1:
            row["apply"] = st.checkbox(
                label + ("" if supported else "  (not wired)"),
                value=bool(row["apply"]),
                key=f"apply_{key}",
                disabled=not supported,
            )
        with c2:
            row["fail"] = float(st.number_input(
                f"Fail ({unit})",
                value=float(row["fail"]),
                step=(0.001 if unit == "dB" else 1.0),
                format=("%.3f" if unit == "dB" else "%.4f"),
                key=f"fail_{key}",
                label_visibility="collapsed",
                disabled=not supported or not row["apply"],
            ))

# Read overrides for the engine
otdr = st.session_state.otdr_settings


def _override(key, default):
    row = otdr.get(key) or {}
    return float(row["fail"]) if row.get("apply") else default


THRESHOLDS = {
    "REBURN_THRESHOLD":     _override("bidir_splice_loss",
                                        float(engine.REBURN_THRESHOLD)),
    "SINGLE_DIR_THRESHOLD": _override("unidir_splice_loss",
                                        float(engine.SINGLE_DIR_THRESHOLD)),
    "BIDIR_CONNECTOR_LOSS": _override("bidir_connector_loss",
                                        float(engine.BIDIR_CONNECTOR_LOSS)),
    "LAUNCH_BAD_REFL_DB":   _override("reflectance",
                                        float(engine.LAUNCH_BAD_REFL_DB)),
}


# ─────────────────────────────────────────────────────────────────────────────
#  Step 1 — pick folders
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("1. Pick folders")

c1, c2 = st.columns(2)
with c1:
    st.markdown("**A-direction folder**")
    picked_a = _pick_folder("Browse for A…", key="browse_a")
    if picked_a:
        st.session_state.dir_a = picked_a
    dir_a = st.text_input(
        "A path",
        value=st.session_state.get("dir_a", ""),
        key="dir_a_input",
        help="Folder containing the A-direction .sor / .json files.",
    )
with c2:
    st.markdown("**B-direction folder**")
    picked_b = _pick_folder("Browse for B…", key="browse_b")
    if picked_b:
        st.session_state.dir_b = picked_b
    dir_b = st.text_input(
        "B path",
        value=st.session_state.get("dir_b", ""),
        key="dir_b_input",
        help="Folder containing the B-direction .sor / .json files.",
    )

# Inventory + content-sniff guard
inv_a = _walk_otdr(dir_a) if dir_a and os.path.isdir(dir_a) else []
inv_b = _walk_otdr(dir_b) if dir_b and os.path.isdir(dir_b) else []

c1, c2 = st.columns(2)
with c1:
    if dir_a:
        if not os.path.isdir(dir_a):
            st.error(f"Not a folder: {dir_a}")
        elif not inv_a:
            st.error("No valid .sor / .json OTDR files found in A.")
        else:
            st.success(f"A: {len(inv_a)} valid OTDR file(s).")
with c2:
    if dir_b:
        if not os.path.isdir(dir_b):
            st.error(f"Not a folder: {dir_b}")
        elif not inv_b:
            st.error("No valid .sor / .json OTDR files found in B.")
        else:
            st.success(f"B: {len(inv_b)} valid OTDR file(s).")


# ─────────────────────────────────────────────────────────────────────────────
#  Step 2 — run
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("2. Run")
run = st.button(
    "Generate report",
    type="primary",
    use_container_width=True,
    disabled=not (inv_a and inv_b),
)


def _apply_overrides(thresholds):
    saved = {k: getattr(engine, k) for k in thresholds if hasattr(engine, k)}
    for k, v in thresholds.items():
        if hasattr(engine, k):
            setattr(engine, k, v)
    return saved


def _restore_overrides(saved):
    for k, v in saved.items():
        setattr(engine, k, v)


if run:
    prog = st.progress(0.0, text="Staging A files…")
    staged_a, warns_a = _stage_flat(inv_a, "splice_desk_a_")
    for w in warns_a:
        st.warning(w)
    prog.progress(0.10, text="Staging B files…")
    staged_b, warns_b = _stage_flat(inv_b, "splice_desk_b_")
    for w in warns_b:
        st.warning(w)

    saved = _apply_overrides(THRESHOLDS)
    log_buf = io.StringIO()
    try:
        prog.progress(0.25, text="Loading OTDR files…")
        with redirect_stdout(log_buf):
            fa, fb = load_all(staged_a, staged_b)
        if not fa:
            st.error("Zero fibers loaded from A.")
            st.stop()

        prog.progress(0.32, text="Normalizing untrimmed events…")
        with redirect_stdout(log_buf):
            for r in list(fa.values()) + list(fb.values()):
                r["_raw_events"] = r["events"]
                r["events"] = _normalize_untrimmed_events(r["events"])

        prog.progress(0.40, text="Discovering closures…")
        with redirect_stdout(log_buf):
            cand = discover_splices(fa)
            splices = refine_closure_centers(fa, cand)

        n_fibers = max(fa.keys())
        span_km = max(
            (r["events"][-1]["dist_km"] for r in fa.values() if r.get("events")),
            default=0.0,
        )

        prog.progress(0.50, text="Detecting launch issues…")
        with redirect_stdout(log_buf):
            launch_a, launch_b = detect_launch_issues(fa, fb, span_km=span_km)

        prog.progress(0.60, text="Pass 1 — bidirectional analysis…")
        with redirect_stdout(log_buf):
            results = analyze_all(fa, fb, splices,
                                   threshold=THRESHOLDS["REBURN_THRESHOLD"],
                                   total_span_a=span_km)
            a_standalone = scan_a_standalone_events(
                fa, splices, results, span_km, fibers_b=fb)
            seen = {**results, **a_standalone}
            bfill = scan_b_past_breaks(fa, fb, splices,
                                        THRESHOLDS["REBURN_THRESHOLD"],
                                        seen, span_km)
            seen = {**seen, **bfill}
            b_side = scan_b_side_breaks(fa, fb, splices, seen, span_km)
            all_results = {**results, **a_standalone, **bfill, **b_side}

        prog.progress(0.75, text="Annotating gainers + connector losses…")
        with redirect_stdout(log_buf):
            apply_field_gainer_rule(all_results, span_km)
            apply_connector_loss_rule(
                all_results, threshold=THRESHOLDS["BIDIR_CONNECTOR_LOSS"])
            all_results, splices = split_offsplice_events_into_own_columns(
                all_results, splices, total_span_km=span_km)

        prog.progress(0.88, text="Building ribbon grid + xlsx…")
        # Output filename derived from inputs + write next to A folder.
        out_dir = os.path.join(os.path.dirname(os.path.abspath(dir_a)),
                                "splice_report_output")
        os.makedirs(out_dir, exist_ok=True)
        a_name = os.path.basename(os.path.normpath(dir_a)) or "A"
        b_name = os.path.basename(os.path.normpath(dir_b)) or "B"
        xlsx_path = os.path.join(out_dir, f"splice_report_{a_name}_to_{b_name}.xlsx")

        with redirect_stdout(log_buf):
            cells, la, lb = build_ribbon_data(
                all_results, n_fibers, int(engine.RIBBON_SIZE),
                len(splices),
                launch_issues={"A": launch_a, "B": launch_b},
            )
            write_xlsx(cells, splices, n_fibers, int(engine.RIBBON_SIZE),
                        xlsx_path, launch_cells_a=la, launch_cells_b=lb,
                        splices_meta=splices,
                        site_a=a_name, site_b=b_name, span_km=span_km)
        prog.progress(1.0, text="Done.")
    finally:
        _restore_overrides(saved)
        shutil.rmtree(staged_a, ignore_errors=True)
        shutil.rmtree(staged_b, ignore_errors=True)

    st.success(f"Report written: {xlsx_path}")
    with open(xlsx_path, "rb") as fh:
        st.download_button(
            "Download Excel report",
            data=fh.read(),
            file_name=os.path.basename(xlsx_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with st.expander("Engine log"):
        st.code(log_buf.getvalue() or "(no log output)", language="text")
