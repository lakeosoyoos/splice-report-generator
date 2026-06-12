# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for the Splice Report desktop app (Windows).
#
# CRITICAL TOOLCHAIN — DO NOT CHANGE WITHOUT READING THIS FIRST:
#   * Build with Python 3.11 (NOT 3.12 or newer).  We pin setuptools 65.5.1
#     below; that version's pkg_resources uses pkgutil.ImpImporter, which
#     Python 3.12 REMOVED.  On 3.12 the packaged exe crashes at launch with
#     "module 'pkgutil' has no attribute 'ImpImporter'".
#   * setuptools must be EXACTLY 65.5.1, installed LAST in
#     requirements-desktop.txt.  Newer setuptools makes pkg_resources strict
#     and crashes the packaged exe with "InvalidVersion: '.../SpliceReport'".
#   * pkg_resources' vendored jaraco / packaging / platformdirs / appdirs /
#     more_itertools packages are bundled THREE WAYS:
#         (a) collect_submodules("pkg_resources") + collect_submodules(
#             "setuptools") + collect_data_files("pkg_resources")
#         (b) installed as REAL top-level packages by requirements-desktop.txt
#         (c) collect_all() in a try/except for each, plus listed in
#             hiddenimports as belt-and-suspenders.
#
# A green PyInstaller build tells you NOTHING about whether the exe boots.
# Every one of our broken builds compiled cleanly.  The ONLY proof is the
# CI boot self-test in .github/workflows/build-windows.yml.  Treat a green
# build with a missing or failing boot test as broken.

import os
import sys
from PyInstaller.utils.hooks import (
    collect_all, collect_submodules, collect_data_files,
)

APP_NAME = "SpliceReport"
SPEC_DIR = os.path.dirname(os.path.abspath(SPEC))
REPO_ROOT = os.path.dirname(SPEC_DIR)

block_cipher = None


# ─── Heavy shells we want to fully bundle ─────────────────────────────
_to_collect = ["streamlit", "altair", "numpy", "openpyxl", "reportlab"]
# pyarrow / pandas are not direct deps; skip unless they sneak in via
# streamlit transitives — collect_all handles missing ones gracefully.
_optional = ["pyarrow", "pandas", "matplotlib", "scipy"]

datas, binaries, hiddenimports = [], [], []

for name in _to_collect + _optional:
    try:
        d, b, h = collect_all(name)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception as e:
        # An optional package that isn't installed → ignore.
        print(f"[spec] skip collect_all({name}): {e}")


# ─── (a) pkg_resources + setuptools — submodules + data ───────────────
hiddenimports += collect_submodules("pkg_resources")
hiddenimports += collect_submodules("setuptools")
datas += collect_data_files("pkg_resources")


# ─── (c) collect_all the vendored packages even though they're also
#         installed as top-level (b in requirements).  Belt-and-suspenders.
for name in ("jaraco.text", "jaraco.functools", "jaraco.context",
             "more_itertools", "packaging", "platformdirs", "appdirs",
             "ordered_set"):
    try:
        d, b, h = collect_all(name)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception as e:
        print(f"[spec] skip collect_all({name}): {e}")


# ─── Explicit hidden imports ──────────────────────────────────────────
hiddenimports += [
    # Engine and UI live next to the launcher; the runtime sys.path tweak
    # is enough to find them, but listing them here makes PyInstaller pull
    # them in as Python modules so they're always available.
    "splicereportmatchexfo",
    "components.otdr_settings",
    # Folder picker — wrapped in try/except in desktop_app.py but Tk and
    # filedialog can be missed by PyInstaller's auto-detection on
    # Windows.
    "tkinter",
    "tkinter.filedialog",
    # Streamlit internals that get loaded via dynamic dispatch.
    "streamlit.web.cli",
    "streamlit.runtime",
    "streamlit.runtime.scriptrunner.magic_funcs",
]


# ─── Bundle our own .py + the custom HTML component ───────────────────
# Engine
datas += [(os.path.join(REPO_ROOT, "splicereportmatchexfo.py"), ".")]
# Custom component (the Streamlit component_v1 declare_component looks
# for the HTML next to the __init__.py)
datas += [(os.path.join(REPO_ROOT, "components", "otdr_settings",
                         "__init__.py"),
            "components/otdr_settings")]
datas += [(os.path.join(REPO_ROOT, "components", "otdr_settings",
                         "index.html"),
            "components/otdr_settings")]
# Desktop UI (loaded by the launcher when no fresh download is available)
datas += [(os.path.join(SPEC_DIR, "desktop_app.py"), "desktop")]


# ─── Excludes — large native-lib PDF engines we don't use ────────────
excludes = ["weasyprint", "cairocffi", "pango", "gobject", "PyQt5", "PyQt6",
             "PySide2", "PySide6"]


a = Analysis(
    [os.path.join(SPEC_DIR, "launcher.py")],
    pathex=[REPO_ROOT, SPEC_DIR],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)


exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,           # UPX corrupts some PyInstaller bootloaders on Windows.
    console=False,       # WINDOWED — no command-prompt window pops up.
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)


coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)
