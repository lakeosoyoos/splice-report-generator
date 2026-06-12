"""
Splice Report — PyInstaller launcher
====================================
This script is the entry point of the frozen Windows .exe.  It:

  1. Redirects stdout/stderr to a log file (a frozen windowed app's
     stdout is None and any print() crashes the process).
  2. Pre-seeds Streamlit's first-run prompt + headless env vars so the
     hidden subprocess doesn't block on stdin.
  3. Optionally auto-updates the engine + desktop UI files from
     raw.githubusercontent.com / main into ~/.spliceReport/engine and
     validates the download all-or-nothing (non-empty + contains
     b"def " + compile() + import smoke-check).  On any failure, falls
     back to the bundled copies that ship inside the .exe.
  4. Spawns a background thread that waits until the Streamlit server
     answers /_stcore/health = "ok", then opens the user's browser.
  5. Detects an already-running instance and just opens a new tab into
     it rather than starting a dead second server.
  6. Boots Streamlit via its CLI entry point so all of its plumbing
     (websocket, static files, dev tooling) just works.

NOTE — boundary call-out for future-me:
The auto-update mechanism runs *after* the bundle bootstraps.  It can
therefore only ship changes to the engine + UI .py files; it can never
fix a bundle that won't boot at all.  Any change to launcher.py, the
.spec file, Python version, or bundled deps requires a fresh download
of the build from the GitHub Release.
"""
from __future__ import annotations

import os
import sys
import json
import time
import threading
import socket
import shutil
import subprocess
import urllib.request
import urllib.error
import webbrowser
from pathlib import Path


APP_NAME       = "SpliceReport"
APP_DIR_NAME   = ".spliceReport"
HOST           = "127.0.0.1"
PORT           = 8501
HEALTH_URL     = f"http://{HOST}:{PORT}/_stcore/health"
APP_URL        = f"http://{HOST}:{PORT}"
GH_OWNER       = "lakeosoyoos"
GH_REPO        = "splice-report-generator"
GH_BRANCH      = "main"
RAW_URL_FMT    = ("https://raw.githubusercontent.com/"
                  f"{GH_OWNER}/{GH_REPO}/{GH_BRANCH}/{{path}}")

# Files we fetch from main on each launch.  Order matters for the
# import smoke check: engine module must compile before the UI that
# imports it.
ENGINE_FILES = [
    "splicereportmatchexfo.py",
    "components/otdr_settings/__init__.py",
    "components/otdr_settings/index.html",
    "desktop/desktop_app.py",
]


# ─────────────────────────────────────────────────────────────────────────────
#  1. Redirect stdout / stderr (frozen windowed exe has None for both)
# ─────────────────────────────────────────────────────────────────────────────
def _redirect_output_to_log() -> Path:
    log_dir = Path.home() / APP_DIR_NAME
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"{APP_NAME.lower()}.log"
    # Open unbuffered so a crash flushes immediately.
    fh = open(log_path, "a", buffering=1, encoding="utf-8", errors="replace")
    sys.stdout = fh
    sys.stderr = fh
    print(f"\n=== {APP_NAME} launch {time.strftime('%Y-%m-%d %H:%M:%S')} ===")
    print(f"frozen={getattr(sys, 'frozen', False)}  executable={sys.executable}")
    return log_path


# ─────────────────────────────────────────────────────────────────────────────
#  2. Silence Streamlit's first-run email prompt (blocks on stdin)
# ─────────────────────────────────────────────────────────────────────────────
def _silence_first_run_prompt() -> None:
    cred_dir = Path.home() / ".streamlit"
    cred_dir.mkdir(parents=True, exist_ok=True)
    cred_path = cred_dir / "credentials.toml"
    if not cred_path.exists():
        cred_path.write_text('[general]\nemail = ""\n', encoding="utf-8")
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "true")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_GLOBAL_DEVELOPMENT_MODE", "false")
    os.environ.setdefault("STREAMLIT_SERVER_ADDRESS", HOST)
    os.environ.setdefault("STREAMLIT_SERVER_PORT", str(PORT))


# ─────────────────────────────────────────────────────────────────────────────
#  3. Auto-update engine + UI files (all-or-nothing)
# ─────────────────────────────────────────────────────────────────────────────
def _bundled_dir() -> Path:
    """Where the bundled copies of engine/UI files live inside the .exe."""
    if getattr(sys, "frozen", False):
        # PyInstaller one-folder: sys._MEIPASS for onefile,
        # os.path.dirname(sys.executable) for one-folder.
        return Path(getattr(sys, "_MEIPASS",
                            os.path.dirname(sys.executable)))
    return Path(__file__).resolve().parent.parent  # repo root in dev


def _fetch(url: str, timeout: int = 15) -> bytes | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": APP_NAME})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status != 200:
                return None
            return resp.read()
    except (urllib.error.URLError, socket.timeout, ConnectionError):
        return None


def _validate_py(data: bytes, rel_path: str) -> bool:
    """Cheap sanity check: non-empty + has at least one 'def ' + compiles.
    HTML files only need to be non-empty."""
    if not data:
        return False
    if rel_path.endswith(".py"):
        if b"def " not in data:
            return False
        try:
            compile(data, rel_path, "exec")
        except SyntaxError:
            return False
    return True


def _try_auto_update(staging: Path) -> bool:
    """Download every ENGINE_FILE into a staging dir.  Return True only if
    every file fetched + validated.  On any failure, the staging dir is
    discarded and the caller falls back to the bundled copies."""
    if staging.exists():
        shutil.rmtree(staging, ignore_errors=True)
    staging.mkdir(parents=True, exist_ok=True)
    for rel in ENGINE_FILES:
        url = RAW_URL_FMT.format(path=rel)
        data = _fetch(url)
        if data is None:
            print(f"auto-update: fetch failed for {rel}")
            return False
        if not _validate_py(data, rel):
            print(f"auto-update: validation failed for {rel}")
            return False
        target = staging / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
    return True


def _prepare_engine() -> tuple[Path, str]:
    """Figure out which engine source we'll run.  Returns (engine_dir,
    source_label) where engine_dir is added to sys.path before Streamlit
    boots the UI script, and source_label appears in the sidebar."""
    bundle = _bundled_dir()
    staging = Path.home() / APP_DIR_NAME / "engine"
    print(f"auto-update: attempting fetch from {GH_OWNER}/{GH_REPO}@{GH_BRANCH}")
    if _try_auto_update(staging):
        print(f"auto-update: ok — using {staging}")
        return staging, "latest (auto-updated)"
    print(f"auto-update: falling back to bundled engine at {bundle}")
    return bundle, "bundled (offline)"


# ─────────────────────────────────────────────────────────────────────────────
#  4. Browser opener (poll /_stcore/health first)
# ─────────────────────────────────────────────────────────────────────────────
def _health_ok() -> bool:
    try:
        with urllib.request.urlopen(HEALTH_URL, timeout=2) as resp:
            return resp.status == 200 and resp.read().strip() == b"ok"
    except (urllib.error.URLError, socket.timeout, ConnectionError):
        return False


def _open_browser_when_ready() -> None:
    """Poll the Streamlit health endpoint, then open the user's browser.
    A fixed delay isn't enough — cold launches can take 20+ seconds while
    PyInstaller unpacks dependencies, and opening the URL too early shows
    a 'connection refused' page that confuses techs into closing it."""
    deadline = time.time() + 90
    while time.time() < deadline:
        if _health_ok():
            try:
                webbrowser.open(APP_URL)
            except Exception as exc:
                print(f"webbrowser.open failed: {exc}")
            return
        time.sleep(0.5)
    print("browser opener: server never returned ok within 90s")


# ─────────────────────────────────────────────────────────────────────────────
#  5. Already-running check (avoid second dead server on double double-click)
# ─────────────────────────────────────────────────────────────────────────────
def _already_running() -> bool:
    return _health_ok()


# ─────────────────────────────────────────────────────────────────────────────
#  6. Main — boot Streamlit
# ─────────────────────────────────────────────────────────────────────────────
def main() -> int:
    log_path = _redirect_output_to_log()
    _silence_first_run_prompt()

    if _already_running():
        print("Another instance is already serving — opening new tab.")
        try:
            webbrowser.open(APP_URL)
        except Exception:
            pass
        return 0

    engine_dir, source_label = _prepare_engine()
    os.environ["SS_ENGINE_DIR"]    = str(engine_dir)
    os.environ["SS_ENGINE_SOURCE"] = source_label

    # The UI script we're going to run.  Prefer the freshly-downloaded
    # desktop_app.py if it's there; otherwise use the bundled copy.
    candidate = engine_dir / "desktop" / "desktop_app.py"
    if candidate.exists():
        ui_script = str(candidate)
    else:
        bundle_ui = _bundled_dir() / "desktop" / "desktop_app.py"
        if not bundle_ui.exists():
            # Last resort: same directory as the launcher (dev mode)
            bundle_ui = Path(__file__).resolve().parent / "desktop_app.py"
        ui_script = str(bundle_ui)
    print(f"UI script: {ui_script}")

    # Kick off the browser opener BEFORE handing control to Streamlit.
    threading.Thread(target=_open_browser_when_ready, daemon=True).start()

    # Boot Streamlit's CLI in-process.  This is the supported entry point
    # for embedding the server.
    from streamlit.web import cli as stcli
    sys.argv = [
        "streamlit", "run", ui_script,
        f"--server.headless=true",
        f"--server.port={PORT}",
        f"--server.address={HOST}",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
    ]
    try:
        return stcli.main()
    except SystemExit as exc:
        # streamlit CLI calls sys.exit() — propagate the code unchanged.
        code = exc.code if isinstance(exc.code, int) else 0
        return code


if __name__ == "__main__":
    sys.exit(main() or 0)
