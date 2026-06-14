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
    and binds to 127.0.0.1:8503 only (Secret Sauce holds 8501).
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
#  Custom EXFO-styled OTDR panel — safety net then import
# ─────────────────────────────────────────────────────────────────────────────
# The component's __init__.py resolves `index.html` via os.path.dirname(__file__).
# In dev mode that's the repo's components/ folder; in the frozen exe it's
# either ~/.spliceReport/engine/components/otdr_settings/ (auto-update path)
# or the PyInstaller _MEIPASS bundle dir (fallback).
#
# The realistic failure mode is "auto-update wrote __init__.py but the HTML
# fetch failed for index.html" — the iframe then loads a 404 and the panel
# goes blank with no error.  This safety net copies the HTML from a known-
# good bundled location into the same directory as __init__.py before the
# component is declared, so declare_component() always sees an HTML file.
def _ensure_component_html() -> None:
    try:
        import components.otdr_settings as _comp  # noqa: F401
        comp_dir = Path(_comp.__file__).resolve().parent
    except Exception:
        return  # import failure surfaces normally below
    expected_html = comp_dir / "index.html"
    if expected_html.exists() and expected_html.stat().st_size > 0:
        return
    # Search the bundled / repo locations for a fallback copy.
    candidates = [
        Path(_HERE).parent / "components" / "otdr_settings" / "index.html",
        Path(_REPO_ROOT) / "components" / "otdr_settings" / "index.html",
    ]
    # PyInstaller _MEIPASS (when frozen)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "components" / "otdr_settings"
                          / "index.html")
    for src in candidates:
        if src.exists() and src.stat().st_size > 0:
            try:
                expected_html.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, expected_html)
                return
            except OSError:
                continue
    # If we couldn't recover, the panel will render blank — let the user
    # know via a visible error rather than a silent empty space.
    st.error(
        f"OTDR settings panel HTML is missing at {expected_html}. "
        "The desktop app couldn't auto-download it and no bundled copy "
        "was reachable.  Restart with internet to retry, or reinstall."
    )


_ensure_component_html()
from components.otdr_settings import otdr_settings as otdr_settings_component  # noqa: E402


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
    "Pick a folder OR a zip for each direction, click Run, and the report "
    "drops next to the inputs."
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


def _is_zip_path(path: str) -> bool:
    return bool(path) and path.lower().endswith(".zip") and os.path.isfile(path)


def _walk_zip_otdr(zip_path: str) -> list[str]:
    """Walk inside a zip and return the names of entries that pass the
    content-sniff guard.  Names are returned as 'zip:<entry>' so the
    staging step can pull them back out.  Skips macOS metadata (__MACOSX,
    .DS_Store)."""
    out = []
    try:
        with zipfile.ZipFile(zip_path) as zf:
            for nm in zf.namelist():
                if nm.endswith("/") or nm.startswith("__MACOSX") or "/.DS_Store" in nm:
                    continue
                low = nm.lower()
                if not (low.endswith(".sor") or low.endswith(".json")):
                    continue
                try:
                    with zf.open(nm) as fh:
                        head = fh.read(4096)
                except (KeyError, RuntimeError):
                    continue
                if low.endswith(".sor"):
                    if SOR_MAGIC in head[:8]:
                        out.append(f"zip:{nm}")
                else:
                    if any(k in head for k in JSON_KEYS):
                        out.append(f"zip:{nm}")
    except (zipfile.BadZipFile, OSError):
        return []
    return out


def _walk_otdr(path: str) -> list[str]:
    """Dispatch on whether the input is a folder or a zip.  Returns a list
    of opaque tokens that _stage_flat() knows how to materialize:
      - folder mode: absolute filesystem paths
      - zip mode:    'zip:<entry-name-inside-zip>'  (paired with the zip
                     path in _stage_flat via a closure)"""
    if not path:
        return []
    if _is_zip_path(path):
        return _walk_zip_otdr(path)
    if not os.path.isdir(path):
        return []
    out = []
    for root, _, files in os.walk(path):
        for f in files:
            full = os.path.join(root, f)
            if _is_otdr_file(full):
                out.append(full)
    return out


def _stage_flat(paths: list[str], prefix: str,
                  zip_source: str | None = None) -> tuple[str, list[str]]:
    """Materialize each input into a fresh temp dir with a flat layout.
    De-duplicates basenames by suffixing __<n>.  Handles both filesystem
    paths and 'zip:<entry>' tokens — when a token has the zip: prefix,
    the entry is extracted from zip_source.
    Returns (temp_dir, [warnings])."""
    tmp = tempfile.mkdtemp(prefix=prefix)
    warns: list[str] = []
    seen: dict[str, int] = {}

    def _next_dest(base: str) -> str:
        stem, ext = os.path.splitext(base)
        n = seen.get(base, 0)
        seen[base] = n + 1
        if n == 0:
            return base
        nm = f"{stem}__{n}{ext}"
        warns.append(f"Duplicate basename '{base}' — staged as '{nm}'.")
        return nm

    zf = None
    try:
        if zip_source is not None:
            try:
                zf = zipfile.ZipFile(zip_source)
            except (zipfile.BadZipFile, OSError) as exc:
                warns.append(f"Could not open zip '{zip_source}' ({exc}).")
                return tmp, warns
        for src in paths:
            if src.startswith("zip:"):
                if zf is None:
                    warns.append(f"Skipped '{src}' — no zip source available.")
                    continue
                entry = src[4:]
                base = os.path.basename(entry) or entry
                dest_name = _next_dest(base)
                try:
                    with zf.open(entry) as srcfh, \
                            open(os.path.join(tmp, dest_name), "wb") as dstfh:
                        shutil.copyfileobj(srcfh, dstfh)
                except (KeyError, OSError) as exc:
                    warns.append(f"Skipped '{src}' ({exc}).")
            else:
                base = os.path.basename(src)
                dest_name = _next_dest(base)
                try:
                    shutil.copy2(src, os.path.join(tmp, dest_name))
                except OSError as exc:
                    warns.append(f"Skipped '{src}' ({exc}).")
    finally:
        if zf is not None:
            zf.close()
    return tmp, warns


def _pick_folder(button_label: str, key: str) -> str | None:
    """Open the native folder picker and return the chosen path (or None).
    Uses tkinter — no extra deps."""
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


def _pick_zip(button_label: str, key: str) -> str | None:
    """Open a native file picker filtered to .zip files."""
    if st.button(button_label, key=key):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            path = filedialog.askopenfilename(
                parent=root,
                filetypes=[("Zip archive", "*.zip"), ("All files", "*")],
            )
            root.destroy()
            return path or None
        except Exception as exc:
            st.warning(f"Native file picker unavailable ({exc}). "
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

    # Widen the sidebar so the EXFO-styled table fits cleanly, same as web.
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

    st.markdown("---")
    # ── EXFO-styled OTDR settings panel (same component as the web app) ──
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
    # Component key encodes the active profile so switching customers
    # forces a re-mount with the new initial values.
    _commit = otdr_settings_component(
        _otdr_rows_for_component,
        default=None,
        key=f"otdr_component::{st.session_state.otdr_profile}",
    )
    if _commit:
        for key, vals in _commit.items():
            st.session_state.otdr_settings[key] = {
                "apply":   bool(vals.get("apply")),
                "fail":    float(vals.get("fail", 0.0)),
                "warning": float(vals.get("warning", 0.0)),
            }

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
#  Step 1 — pick folders or zips
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("1. Pick A and B")

# A returning picked path must be written to the TEXT_INPUT's own widget
# key, not a separate session_state slot.  Streamlit's widget-state takes
# precedence over `value=` on re-render, so writing to `dir_a` (separate
# slot) leaves the text_input stuck on its prior empty value and the
# picked path never appears.  After we mutate the widget key we call
# st.rerun() so the visible box refreshes immediately.
c1, c2 = st.columns(2)
with c1:
    st.markdown("**A-direction**")
    b1, b2 = st.columns(2)
    with b1:
        picked = _pick_folder("Browse folder…", key="browse_a_folder")
        if picked:
            st.session_state["dir_a_input"] = picked
            st.rerun()
    with b2:
        picked = _pick_zip("Browse zip…", key="browse_a_zip")
        if picked:
            st.session_state["dir_a_input"] = picked
            st.rerun()
    dir_a = st.text_input(
        "A path",
        key="dir_a_input",
        help="Folder OR .zip containing the A-direction .sor / .json files.",
    )
with c2:
    st.markdown("**B-direction**")
    b1, b2 = st.columns(2)
    with b1:
        picked = _pick_folder("Browse folder…", key="browse_b_folder")
        if picked:
            st.session_state["dir_b_input"] = picked
            st.rerun()
    with b2:
        picked = _pick_zip("Browse zip…", key="browse_b_zip")
        if picked:
            st.session_state["dir_b_input"] = picked
            st.rerun()
    dir_b = st.text_input(
        "B path",
        key="dir_b_input",
        help="Folder OR .zip containing the B-direction .sor / .json files.",
    )


def _path_status(label: str, path: str) -> tuple[list[str], str | None]:
    """Returns (inventory_tokens, error_message_or_None)."""
    if not path:
        return [], None
    if _is_zip_path(path):
        items = _walk_otdr(path)
        if not items:
            return [], f"No valid .sor / .json OTDR files found in {label} zip."
        return items, None
    if not os.path.isdir(path):
        return [], f"Not a folder or zip: {path}"
    items = _walk_otdr(path)
    if not items:
        return [], f"No valid .sor / .json OTDR files found in {label}."
    return items, None


# Inventory + content-sniff guard (handles folder OR zip)
inv_a, err_a = _path_status("A", dir_a)
inv_b, err_b = _path_status("B", dir_b)

c1, c2 = st.columns(2)
with c1:
    if dir_a:
        if err_a:
            st.error(err_a)
        else:
            kind = "zip" if _is_zip_path(dir_a) else "folder"
            st.success(f"A: {len(inv_a)} valid OTDR file(s) in {kind}.")
with c2:
    if dir_b:
        if err_b:
            st.error(err_b)
        else:
            kind = "zip" if _is_zip_path(dir_b) else "folder"
            st.success(f"B: {len(inv_b)} valid OTDR file(s) in {kind}.")


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
# ── Test hook ─────────────────────────────────────────────────────────
# AppTest's button-click API is flaky for type="primary" buttons, so the
# end-to-end test (desktop/tests/test_e2e_generate.py) sets a magic
# session-state flag instead.  This branch is a no-op in production —
# real users never set that key, and the UI is unchanged.
if st.session_state.get("__test_force_generate__", False):
    run = True
    # Reset so the next AppTest.run() (e.g. a re-render after Generate
    # finishes) doesn't trigger Generate again in a loop.
    st.session_state["__test_force_generate__"] = False


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
    staged_a, warns_a = _stage_flat(
        inv_a, "splice_desk_a_",
        zip_source=dir_a if _is_zip_path(dir_a) else None,
    )
    for w in warns_a:
        st.warning(w)
    prog.progress(0.10, text="Staging B files…")
    staged_b, warns_b = _stage_flat(
        inv_b, "splice_desk_b_",
        zip_source=dir_b if _is_zip_path(dir_b) else None,
    )
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
        # detect_launch_issues returns a SINGLE dict keyed by fiber number
        # ({fnum: launch_issue_dict}), not a per-direction (A, B) tuple —
        # see the same call site in app.py.  The earlier desktop version
        # tried to unpack the dict into two variables, which iterates the
        # dict's keys and raised "too many values to unpack (expected 2)".
        # `span_km` was also being silently swallowed; the engine uses
        # first_splice_km.
        first_splice_km = splices[0]["position_km"] if splices else None
        with redirect_stdout(log_buf):
            launch_issues = detect_launch_issues(
                fa, fb, first_splice_km=first_splice_km)

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
        # Output filename derived from inputs + written next to the A input
        # (whether A is a folder or a zip).  Strip a trailing .zip from the
        # display name so the output filename reads cleanly.
        a_abs = os.path.abspath(dir_a)
        b_abs = os.path.abspath(dir_b)
        a_parent = os.path.dirname(a_abs) if _is_zip_path(dir_a) else os.path.dirname(a_abs)
        out_dir = os.path.join(a_parent, "splice_report_output")
        os.makedirs(out_dir, exist_ok=True)
        a_name = os.path.basename(os.path.normpath(dir_a)) or "A"
        b_name = os.path.basename(os.path.normpath(dir_b)) or "B"
        for suf in (".zip", ".ZIP"):
            if a_name.endswith(suf): a_name = a_name[:-len(suf)]
            if b_name.endswith(suf): b_name = b_name[:-len(suf)]
        # Sanitize before building the filename — free-text site names
        # can contain `/ : * ? < > |` which break browser-download
        # dialogs and Windows file creation.  See bug #4 fix 2026-06-13.
        from splicereportmatchexfo import _safe_filename_part
        a_name_safe = _safe_filename_part(a_name, "A")
        b_name_safe = _safe_filename_part(b_name, "B")
        xlsx_path = os.path.join(
            out_dir, f"splice_report_{a_name_safe}_to_{b_name_safe}.xlsx")

        with redirect_stdout(log_buf):
            cells, la, lb = build_ribbon_data(
                all_results, n_fibers, int(engine.RIBBON_SIZE),
                len(splices),
                launch_issues=launch_issues,
            )
            write_xlsx(cells, splices, n_fibers, int(engine.RIBBON_SIZE),
                        xlsx_path,
                        site_a=a_name, site_b=b_name, span_km=span_km,
                        launch_cells_a=la, launch_cells_b=lb,
                        fibers_a=fa, fibers_b=fb,
                        all_results=all_results)
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
