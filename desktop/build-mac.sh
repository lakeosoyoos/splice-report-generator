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
#  Requires Python 3.11 (NOT 3.12 or newer — see SpliceReport.spec for why).
#  Uses the portable Python interpreter at desktop/.python311/python/bin/
#  python3.11 that the install script left behind; falls back to a system
#  python3.11 if that's missing.
# =============================================================================
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$HERE"

# ── 1. Find Python 3.11 ──────────────────────────────────────────────────────
PORTABLE_PY="$HERE/.python311/python/bin/python3.11"
if [[ -x "$PORTABLE_PY" ]]; then
    PY="$PORTABLE_PY"
elif command -v python3.11 >/dev/null 2>&1; then
    PY="$(command -v python3.11)"
else
    echo "[build-mac] ERROR — Python 3.11 not found." >&2
    echo "             Install the portable interpreter (see README_BUILD.txt)" >&2
    echo "             or install Python 3.11 system-wide and re-run." >&2
    exit 1
fi
echo "[build-mac] Using: $PY"
"$PY" --version

# ── 2. Fresh venv ────────────────────────────────────────────────────────────
VENV="$HERE/.venv-mac"
if [[ -d "$VENV" ]]; then
    echo "[build-mac] Removing old $VENV ..."
    rm -rf "$VENV"
fi
"$PY" -m venv "$VENV"
# shellcheck disable=SC1091
source "$VENV/bin/activate"
python -m pip install --upgrade pip wheel

# ── 3. Install deps + re-pin setuptools 65.5.1 LAST ──────────────────────────
pip install -r requirements-desktop.txt
pip install --force-reinstall "setuptools==65.5.1"

# ── 4. PyInstaller build ─────────────────────────────────────────────────────
rm -rf build dist
pyinstaller SpliceReport-mac.spec --noconfirm --clean

if [[ ! -d "dist/SpliceReport.app" ]]; then
    echo "[build-mac] ERROR — dist/SpliceReport.app missing after PyInstaller." >&2
    exit 1
fi

# ── 5. Refresh the .app on the Desktop ──────────────────────────────────────
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
