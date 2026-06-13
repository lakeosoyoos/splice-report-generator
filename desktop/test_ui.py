"""
Desktop UI smoke test
=====================
Runs the desktop Streamlit script through Streamlit's official AppTest
harness (no browser, no .exe) and asserts the widget-state plumbing
that bit us in commit 224dae0:

  * Setting `session_state["dir_a_input"]` (the text_input's own widget
    key) must cause the visible text_input to display that path.
  * The same goes for `dir_b_input`.
  * Setting a path that doesn't exist on disk must surface an error,
    not a misleading success.

This is what a healthy CI check looks like for a widget-state
regression of the same shape as the "dialog opens but path doesn't
stick" bug — independent of any browser or PyInstaller bundling.  Run:

  python3 -m desktop.test_ui

…or via pytest:

  pytest desktop/test_ui.py -v

The test passes its return code through sys.exit so CI can fail the
job on a regression.
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(HERE))


def _run_smoke() -> int:
    try:
        from streamlit.testing.v1 import AppTest
    except ImportError as exc:
        print(f"[test_ui] FAIL — streamlit.testing.v1 import: {exc}",
              file=sys.stderr)
        return 2

    failures: list[str] = []

    # ─── Test 1: widget-state plumbing ──────────────────────────────
    # The picker writes to session_state[<widget-key>]; the text_input
    # must display that value on the next run.
    print("[test_ui] Test 1: dir_a_input / dir_b_input plumbing")
    at = AppTest.from_file(str(HERE / "desktop_app.py"), default_timeout=30)
    at.session_state["dir_a_input"] = "/tmp/__pretend_a__"
    at.session_state["dir_b_input"] = "/tmp/__pretend_b__"
    at.run()
    if at.exception:
        failures.append(f"  test1 raised: {at.exception}")
    else:
        # Find the text_input by key in the widget tree.
        a_box = next(
            (w for w in at.text_input if w.key == "dir_a_input"), None)
        b_box = next(
            (w for w in at.text_input if w.key == "dir_b_input"), None)
        if a_box is None:
            failures.append("  test1: dir_a_input widget not found")
        elif a_box.value != "/tmp/__pretend_a__":
            failures.append(
                f"  test1: dir_a_input widget shows {a_box.value!r}, "
                f"expected '/tmp/__pretend_a__' — the bug we fixed in "
                f"224dae0 has regressed")
        if b_box is None:
            failures.append("  test1: dir_b_input widget not found")
        elif b_box.value != "/tmp/__pretend_b__":
            failures.append(
                f"  test1: dir_b_input widget shows {b_box.value!r}, "
                f"expected '/tmp/__pretend_b__'")

    # ─── Test 2: invalid path surfaces a clear error ───────────────
    print("[test_ui] Test 2: invalid path → visible error")
    at = AppTest.from_file(str(HERE / "desktop_app.py"), default_timeout=30)
    at.session_state["dir_a_input"] = "/this/path/does/not/exist/anywhere"
    at.run()
    if at.exception:
        failures.append(f"  test2 raised: {at.exception}")
    else:
        error_msgs = [e.value for e in at.error]
        if not any("/this/path/does/not/exist" in m or "Not a" in m
                   for m in error_msgs):
            failures.append(
                f"  test2: expected an st.error mentioning the bad "
                f"path; got errors={error_msgs}")

    # ─── Test 3: real folder with a valid SOR → success ────────────
    print("[test_ui] Test 3: real folder w/ valid SOR → success message")
    # Find any one valid SOR from disk to point the test at.  In CI we
    # checkout the repo but Lagrande↔Durkey isn't checked in, so make a
    # tiny stub file with the SR-4731 'Map\x00' header.  The
    # content-sniff guard only checks the first 8 bytes; we don't need
    # the full file to be a real OTDR shoot.
    with tempfile.TemporaryDirectory(prefix="ui_test_") as td:
        stub = Path(td) / "stub.sor"
        stub.write_bytes(b"Map\x00\x00\x00\x00\x00" + b"x" * 64)
        at = AppTest.from_file(str(HERE / "desktop_app.py"),
                                 default_timeout=30)
        at.session_state["dir_a_input"] = td
        at.run()
        if at.exception:
            failures.append(f"  test3 raised: {at.exception}")
        else:
            success_msgs = [s.value for s in at.success]
            if not any("valid OTDR file" in m for m in success_msgs):
                failures.append(
                    f"  test3: expected an st.success mentioning "
                    f"'valid OTDR file'; got successes={success_msgs}, "
                    f"errors={[e.value for e in at.error]}")

    # ─── Test 4: static check that the picker writes to the right slot ─
    # This is the structural check that catches the EXACT class of bug
    # the tech reported ("dialog opens but path doesn't stick").  The
    # picker callback must write to the text_input's OWN key
    # (dir_a_input / dir_b_input) — not a separate session_state slot —
    # because otherwise the widget-state precedence rule leaves the
    # visible box empty.  See feedback_streamlit_widget_state.md.
    print("[test_ui] Test 4: picker writes to widget key, not separate slot")
    import re
    src = (HERE / "desktop_app.py").read_text()
    # Strip comment lines so we don't match commentary about the bug
    # itself (this file's docstring references both keys deliberately).
    code = "\n".join(ln for ln in src.splitlines()
                    if not ln.strip().startswith("#"))
    # Bad pattern — picker writes to `session_state.dir_a` (separate
    # slot, never read by the widget).
    bad_a = re.search(r"session_state\.dir_a\s*=\s*picked", code)
    bad_b = re.search(r"session_state\.dir_b\s*=\s*picked", code)
    if bad_a or bad_b:
        which = [n for n, m in (('dir_a', bad_a), ('dir_b', bad_b)) if m]
        failures.append(
            f"  test4: picker writes to separate session_state slot "
            f"for {which!r} — must write to '<key>_input' instead "
            f"(see feedback_streamlit_widget_state.md)")
    # Good pattern — picker writes to the widget's own key.
    good_a = re.search(
        r"session_state\[[\"']dir_a_input[\"']\]\s*=\s*picked", code)
    good_b = re.search(
        r"session_state\[[\"']dir_b_input[\"']\]\s*=\s*picked", code)
    if not good_a:
        failures.append(
            "  test4: no 'session_state[\"dir_a_input\"] = picked' "
            "found — picker may not populate the A text_input")
    if not good_b:
        failures.append(
            "  test4: no 'session_state[\"dir_b_input\"] = picked' "
            "found — picker may not populate the B text_input")

    if failures:
        print("\n[test_ui] FAIL")
        for line in failures:
            print(line)
        return 1
    print("\n[test_ui] OK — all UI smoke checks passed")
    return 0


# ── pytest entry points (so `pytest desktop/test_ui.py` works too) ──
def test_widget_state_plumbing():
    assert _run_smoke() == 0, "Desktop UI smoke checks failed"


if __name__ == "__main__":
    sys.exit(_run_smoke())
