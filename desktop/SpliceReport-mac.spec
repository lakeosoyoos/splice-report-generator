# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for the Splice Report desktop app (macOS).
# IDENTICAL to SpliceReport.spec (Windows) except for the trailing
# BUNDLE() step that turns the COLLECT directory into a real
# double-clickable .app.  Read the comment block at the top of
# SpliceReport.spec for the toolchain pins and the reason every one
# matters.
#
# This Mac build is for local de-risking only — it flushes
# OS-independent packaging bugs (missing hiddenimports for our own
# modules, the setuptools 65.5.1 pin, the Streamlit first-run prompt
# hang) before we burn a Windows CI cycle.  A green Mac build does
# NOT prove the Windows app launches: different OS, the host may have
# vendored packages installed incidentally.  The Windows CI BOOT
# SELF-TEST in build-windows.yml remains the only authoritative
# check for what the tech downloads.

import os
from PyInstaller.utils.hooks import (
    collect_all, collect_submodules, collect_data_files,
)

APP_NAME = "SpliceReport"
SPEC_DIR = os.path.dirname(os.path.abspath(SPEC))
REPO_ROOT = os.path.dirname(SPEC_DIR)

block_cipher = None


# ─── Heavy shells we want to fully bundle ─────────────────────────────
_to_collect = ["streamlit", "altair", "numpy", "openpyxl", "reportlab"]
_optional = ["pyarrow", "pandas", "matplotlib", "scipy"]

datas, binaries, hiddenimports = [], [], []

for name in _to_collect + _optional:
    try:
        d, b, h = collect_all(name)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception as e:
        print(f"[spec] skip collect_all({name}): {e}")


# ─── (a) pkg_resources + setuptools — submodules + data ───────────────
hiddenimports += collect_submodules("pkg_resources")
hiddenimports += collect_submodules("setuptools")
datas += collect_data_files("pkg_resources")


# ─── (c) collect_all the vendored packages even though they're also
#         installed as top-level (b in requirements-desktop.txt).
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
    "splicereportmatchexfo",
    "sor_reader324802a",
    "json_reader",
    "components.otdr_settings",
    "tkinter",
    "tkinter.filedialog",
    "streamlit.web.cli",
    "streamlit.runtime",
    "streamlit.runtime.scriptrunner.magic_funcs",
]


# ─── Bundle our own .py + the custom HTML component ───────────────────
datas += [(os.path.join(REPO_ROOT, "splicereportmatchexfo.py"), ".")]
datas += [(os.path.join(REPO_ROOT, "sor_reader324802a.py"), ".")]
datas += [(os.path.join(REPO_ROOT, "json_reader.py"), ".")]
datas += [(os.path.join(REPO_ROOT, "components", "otdr_settings",
                         "__init__.py"),
            "components/otdr_settings")]
datas += [(os.path.join(REPO_ROOT, "components", "otdr_settings",
                         "index.html"),
            "components/otdr_settings")]
datas += [(os.path.join(SPEC_DIR, "desktop_app.py"), "desktop")]


# ─── Excludes ─────────────────────────────────────────────────────────
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
    upx=False,
    console=False,           # WINDOWED — no Terminal popup on launch.
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


# ─── BUNDLE: wrap the COLLECT into a real .app on macOS ──────────────
app = BUNDLE(
    coll,
    name=f"{APP_NAME}.app",
    icon=None,                       # drop a .icns here to brand it later
    bundle_identifier="com.lakeosoyoos.splicereport",
    info_plist={
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": "Splice Report",
        "CFBundleShortVersionString": "1.0.0",
        "CFBundleVersion": "1.0.0",
        "NSHighResolutionCapable": True,
        # Hide from the Dock briefly during the cold launch (a real .app
        # without a window for ~10 s otherwise gets the bouncing icon).
        "LSUIElement": False,
    },
)
