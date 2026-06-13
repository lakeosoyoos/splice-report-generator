Splice Report — desktop build
==============================

This folder packages the Streamlit splice-report app as a downloadable
Windows .exe.  The web app keeps running on Streamlit Cloud unchanged;
this is a parallel distribution channel for techs who want to run big-zip
SOR jobs locally, with no upload limit and no internet round-trip.


Folder contents
---------------
  desktop_app.py            Local Streamlit UI (folder picker version
                            of the web app).
  launcher.py               PyInstaller entry point.  Redirects stdout
                            to a log file, silences Streamlit's first-run
                            prompt, auto-updates the engine + UI .py
                            files from main (with all-or-nothing
                            validation + fallback to the bundled copy),
                            polls the health endpoint, and opens the
                            browser when the server is actually ready.
  SpliceReport.spec         PyInstaller spec.  Read its top-of-file
                            comment block before changing anything — the
                            toolchain has three interlocked pins that
                            crash the exe at launch if any are wrong.
  requirements-desktop.txt  Pinned deps.  setuptools 65.5.1 first, the
                            vendored jaraco/packaging/platformdirs/etc.
                            as real top-level packages, then runtime
                            deps, then pyinstaller.
  build.bat                 Local Windows one-click build.
  README_BUILD.txt          This file.


Toolchain (read before touching anything)
-----------------------------------------
  * Python 3.11 — NOT 3.12 or newer.  3.12 removed pkgutil.ImpImporter
    which pkg_resources at our pinned setuptools version still uses.
  * setuptools 65.5.1 — installed LAST so it wins against transitive
    bumps.  Newer setuptools makes pkg_resources strict and the frozen
    exe crashes at launch with "InvalidVersion".
  * Vendored jaraco/packaging/platformdirs/appdirs/more_itertools/
    ordered_set bundled three ways: collect_submodules in the spec,
    real top-level packages in requirements, and collect_all() in the
    spec.

A green PyInstaller build tells you NOTHING about whether the exe
launches.  Every one of our broken builds compiled cleanly.  The ONLY
proof is the boot self-test in .github/workflows/build-windows.yml —
treat a green build with a missing or failing boot test as broken.


Local build (Windows machine)
-----------------------------
  1. Install Python 3.11 from python.org.
  2. Open a Command Prompt in this folder.
  3. Run:  build.bat
  4. Output:
       dist\SpliceReport\SpliceReport.exe
       dist\SpliceReport-Windows.zip
  Note: local build is NOT verified to launch.  For an authoritative
  pass/fail, push to GitHub and let CI run the boot self-test.


Local Mac build (preview / de-risk only — not for techs)
--------------------------------------------------------
  1. Use the system /usr/bin/python3 (currently 3.9.x).  Same interpreter
     the Secret Sauce Mac app was built with.  Any Python BELOW 3.12
     works because setuptools 65.5.1 needs pkgutil.ImpImporter which 3.12
     removed.  DO NOT install 3.11/3.12 for this build.
  2. From this folder, run:  ./build-mac.sh
     The script installs build deps into ~/Library/Python/3.9/... (via
     --user), runs PyInstaller against SpliceReport-mac.spec, and drops
     SpliceReport.app onto ~/Desktop.
  3. Double-click ~/Desktop/SpliceReport.app.  On a fresh machine
     Gatekeeper says "cannot be opened because it is from an unidentified
     developer" — right-click → Open → Open (one-time), or strip the
     quarantine bit:  xattr -dr com.apple.quarantine ~/Desktop/SpliceReport.app
  This Mac .app is the MAINTAINER's local preview.  Techs get the
  Windows build, and only the Windows CI boot self-test verifies that
  build.  Don't ship the .app.


Dev run (Mac / Linux / Windows — no packaging)
----------------------------------------------
  pip install -r requirements-desktop.txt
  streamlit run desktop/desktop_app.py
  # This runs the UI directly out of source — useful for iterating on
  # the desktop UI without rebuilding the exe.


How a tech installs and uses the app
------------------------------------
  1. Download the zip from the permanent GitHub Release link
     (link is in the release notes; never changes).
  2. Right-click the downloaded zip → Properties → Unblock → OK.
     (Windows tags downloads from the internet; this is what lets the
     extracted exe run.)
  3. Extract All.
  4. Open the extracted folder, double-click SpliceReport.exe.
  5. On first launch, Windows SmartScreen shows "Windows protected your
     PC".  Click "More info" → "Run anyway".  (We don't code-sign the
     exe; signing removes this dialog but is a separate purchase.)
  6. A browser tab opens after a 10-30 second cold-start while the
     bundle unpacks.  No visible window appears during that wait —
     this is normal.
  7. Click "Browse for A..." / "Browse for B...", pick the two
     direction folders, click Generate Report.  The xlsx writes into a
     "splice_report_output" subfolder next to the A folder, and a
     Download button appears.

Each launch checks GitHub for a newer engine + UI and uses it if
download succeeds.  No internet → falls back to the version bundled in
the exe.  Engine fixes you push to main reach techs on their next
double-click — no reinstall needed.  Bundle-level changes (Python
version, new dependency, launcher.py / spec edits) require a fresh
download of the build.
