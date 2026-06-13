#!/usr/bin/env bash
# =============================================================================
#  Splice Report — local macOS build
# =============================================================================
#  Produces dist/SpliceReport.app and refreshes the copy at
#  ~/Desktop/SpliceReport.app so you can double-click it without leaving the
#  repo.  This Mac build is for LOCAL DE-RISKING — it flushes OS-independent
#  packaging bugs before we burn a Windows CI cycle.  A green Mac build does
#  NOT prove the Windows app launches; the Windows CI BOOT SELF-TEST remains
#  the only authoritative check for what the tech downloads.
#
#  PYTHON CHOICE — uses the Mac's built-in /usr/bin/python3 (currently 3.9.x).
#  Same interpreter Secret Sauce was built with.  Any Python BELOW 3.12 works
#  because we pin setuptools==65.5.1, and 3.12 removed pkgutil.ImpImporter
#  which that setuptools relies on.  If you ever build on 3.12+ you'd hit the
#  same ImpImporter crash we already documented on Windows — don't.
#
#  Build deps are installed into the user site (~/Library/Python/3.9/...)
#  via `python3 -m pip install --user`, NOT a venv.  Matches the Secret
#  Sauce pattern and keeps the build reproducible on any Mac.
# =============================================================================
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$HERE"

PY="/usr/bin/python3"
if [[ ! -x "$PY" ]]; then
    echo "[build-mac] ERROR — /usr/bin/python3 missing." >&2
    echo "             macOS ships it by default.  Reinstall the Command Line Tools" >&2
    echo "             with:  xcode-select --install" >&2
    exit 1
fi
PY_VER=$("$PY" --version 2>&1)
echo "[build-mac] Using: $PY ($PY_VER)"

# Guardrail: refuse to build on 3.12+.  See header comment for why.
if "$PY" -c "import sys; sys.exit(0 if sys.version_info < (3, 12) else 1)"; then
    : # ok
else
    echo "[build-mac] ERROR — $PY_VER is 3.12+ which removed pkgutil.ImpImporter." >&2
    echo "             setuptools 65.5.1 (our pinned version) needs it at boot." >&2
    echo "             Use a Python < 3.12, or rebuild the toolchain to use a" >&2
    echo "             newer setuptools.  Don't ship a 3.12 .app — it will crash" >&2
    echo "             at launch the same way the Windows builds did." >&2
    exit 1
fi

# ── 1. Install build deps into the user site (idempotent) ────────────────────
# Same pattern Secret Sauce uses.  --user lands at ~/Library/Python/3.9/...
"$PY" -m pip install --user --upgrade pip wheel >/dev/null
"$PY" -m pip install --user -r requirements-desktop.txt
"$PY" -m pip install --user --force-reinstall "setuptools==65.5.1"

# Put the user-site bin on PATH so the pyinstaller entrypoint resolves.
USER_BIN="$("$PY" -m site --user-base)/bin"
export PATH="$USER_BIN:$PATH"

# ── 2. PyInstaller build ─────────────────────────────────────────────────────
rm -rf build dist
"$PY" -m PyInstaller SpliceReport-mac.spec --noconfirm --clean

if [[ ! -d "dist/SpliceReport.app" ]]; then
    echo "[build-mac] ERROR — dist/SpliceReport.app missing after PyInstaller." >&2
    exit 1
fi

# ── 3. Refresh the .app on the Desktop ──────────────────────────────────────
DEST="$HOME/Desktop/SpliceReport.app"
if [[ -d "$DEST" ]]; then
    echo "[build-mac] Replacing existing $DEST ..."
    rm -rf "$DEST"
fi
cp -R "dist/SpliceReport.app" "$DEST"

# Strip the quarantine bit so Gatekeeper's first-open dialog only nags
# once instead of every relaunch.  This is YOUR Mac, for YOUR testing —
# the warning persists for end users without a signing cert.
xattr -dr com.apple.quarantine "$DEST" 2>/dev/null || true

echo
echo "[build-mac] ============================================================"
echo "[build-mac]  Build OK."
echo "[build-mac]    Source : $HERE/dist/SpliceReport.app"
echo "[build-mac]    Desktop: $DEST"
echo "[build-mac] ============================================================"
echo "[build-mac]  Gatekeeper note for the FIRST run on a fresh machine:"
echo "[build-mac]    right-click SpliceReport.app → Open → Open (one-time)"
echo "[build-mac]    or:  xattr -dr com.apple.quarantine '$DEST'"
echo "[build-mac] ============================================================"
echo "[build-mac]  Reminder — this Mac build is for local preview only."
echo "[build-mac]  Techs get the Windows build, gated by the Windows CI"
echo "[build-mac]  boot self-test.  Don't ship this .app."
echo "[build-mac] ============================================================"
