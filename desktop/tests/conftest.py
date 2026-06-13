"""
Shared scaffolding for the desktop UI test suite.

The other test modules (test_e2e_generate.py — this PR; plus the
threshold/config, input-variant, and engine-edge-case suites that other
agents are building in parallel) all import the symbols below.

Exports
-------
APP_PATH         : pathlib.Path
    Absolute path to desktop/desktop_app.py — the Streamlit script under
    test.  Pass this to ``AppTest.from_file`` or to ``run_streamlit``.

FIXTURE_DIR      : pathlib.Path
    Absolute path to desktop/tests/fixtures/ — the parent directory of
    the per-direction fixture subdirs.

FIXTURE_A_DIR    : pathlib.Path
    Absolute path to the A-direction fixture span (12 SOR files).  Feed
    this to the desktop UI's ``dir_a_input`` text-input slot.

FIXTURE_B_DIR    : pathlib.Path
    Absolute path to the B-direction fixture span (12 SOR files).  Feed
    this to the desktop UI's ``dir_b_input`` text-input slot.

run_streamlit(...)
    Helper that constructs an ``AppTest`` pointed at desktop_app.py with
    sensible defaults — ``default_timeout=60`` (the engine pass takes a
    few seconds on the fixture span), and a ``sys.path`` shim so the
    engine + custom-component modules at the repo root are importable
    from inside the AppTest subprocess.

Conventions for the other agents
--------------------------------
* Always reach for ``run_streamlit`` and the ``FIXTURE_*`` constants
  rather than rebuilding the AppTest plumbing.  If you discover a new
  default you need, add it here so every suite benefits.
* The Generate-report path is triggered via
  ``session_state["__test_force_generate__"] = True`` — see the hook
  near the "Generate report" button in desktop/desktop_app.py.  This
  side-steps AppTest's flaky button-click API for ``type="primary"``
  buttons.
* The desktop UI writes its xlsx into ``splice_report_output/`` next to
  the A input.  Tests are responsible for cleaning that up — see
  test_e2e_generate.py for a pattern.
"""
from __future__ import annotations

import sys
from pathlib import Path

# ── Paths ───────────────────────────────────────────────────────────────
HERE = Path(__file__).resolve().parent
DESKTOP_DIR = HERE.parent
REPO_ROOT = DESKTOP_DIR.parent

APP_PATH: Path = DESKTOP_DIR / "desktop_app.py"
FIXTURE_DIR: Path = HERE / "fixtures"
FIXTURE_A_DIR: Path = FIXTURE_DIR / "span_A"
FIXTURE_B_DIR: Path = FIXTURE_DIR / "span_B"

# Make the engine + custom-component modules at the repo root importable
# from inside AppTest.  AppTest runs the target script in-process, so a
# sys.path edit here propagates to the script.  Also add the tests
# directory itself so sibling test modules can `from conftest import ...`
# (pytest auto-loads conftest as a fixture provider, but does NOT add it
# to sys.path for direct imports).
for p in (REPO_ROOT, DESKTOP_DIR, HERE):
    sp = str(p)
    if sp not in sys.path:
        sys.path.insert(0, sp)


def run_streamlit(default_timeout: float = 60.0, **kwargs):
    """Build an ``AppTest`` pointed at ``desktop_app.py``.

    Parameters
    ----------
    default_timeout : float
        Per-``run()`` timeout in seconds.  60 s is a comfortable upper
        bound for the fixture span (the engine pass is ~3-5 s on a
        modern laptop, but CI runners are slower and AppTest serialises
        widget-tree construction).
    **kwargs
        Forwarded to ``AppTest.from_file``.  Reserved for future use —
        most tests should accept the defaults.

    Returns
    -------
    streamlit.testing.v1.AppTest
        Ready to have ``session_state`` mutated and ``run()`` called.
    """
    from streamlit.testing.v1 import AppTest

    return AppTest.from_file(
        str(APP_PATH), default_timeout=default_timeout, **kwargs
    )


__all__ = [
    "APP_PATH",
    "FIXTURE_DIR",
    "FIXTURE_A_DIR",
    "FIXTURE_B_DIR",
    "run_streamlit",
]
