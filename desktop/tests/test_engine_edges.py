"""
Engine edge-case tests — direct assertions on the engine API.

Hardens the parts of the splicereportmatchexfo pipeline that produce
silently-wrong reports on unusual data.  No UI / AppTest layer here —
tests call the engine functions directly so failures are tight and
fast.

The HIGH/MED/PRE-EXISTING findings these guards correspond to:
  T1  Pre-existing: discover_splices→[] (e.g. <MIN_POP_SPLICE fibers)
       crashes scan_a_standalone_events with TypeError.  [XFAIL]
  T2  H4: orig_loc / term_loc with reserved fs chars break download.
       [XFAIL pending sanitize hook in production]
  T3  H2: apply_field_gainer_rule runs BEFORE off-splice splitter, but
       the splitter ignores is_gainer — gainers stay mis-attributed.
       (current-bug assertion passes; desired-behavior assertion XFAIL)
  T4  H3: Pass-2d B-side breaks must override Pass-2c B-fill, not just
       dead_zone.  (currently drops Pass-2d — XFAIL the desired check.)
  T5  MED: reburn-summary cosmetics on a fully-damaged span.
  T6  MED: mixed JSON+SOR in one dir silently drops one format.  [XFAIL
       pending visible WARN line.]
  T7  MED: UTC calendar-day bucketing splits a PT-evening shoot into
       two days.  [XFAIL pending local-day tolerance.]
  T8  Engine-API signature contract — UI depends on these names.
"""
from __future__ import annotations

import contextlib
import inspect
import io
import shutil
import time
from pathlib import Path

import pytest

from conftest import FIXTURE_A_DIR, FIXTURE_B_DIR  # noqa: F401  (sys.path shim)

import splicereportmatchexfo as engine
from splicereportmatchexfo import (
    MIN_POP_SPLICE,
    apply_field_gainer_rule,
    build_ribbon_data,
    detect_launch_issues,
    discover_splices,
    load_all,
    scan_a_standalone_events,
    scan_b_side_breaks,
    split_offsplice_events_into_own_columns,
    write_xlsx,
)


# ──────────────────────────────────────────────────────────────────────
# T1  empty splice list — pre-existing engine bug guard
# ──────────────────────────────────────────────────────────────────────
def _tiny_a_dir(tmp_path: Path, n: int) -> Path:
    """Copy the first n LAGDUR fixtures into a temp dir."""
    d = tmp_path / "a_small"
    d.mkdir()
    src_files = sorted(p for p in FIXTURE_A_DIR.iterdir()
                       if p.name.lower().endswith(".sor"))
    for p in src_files[:n]:
        shutil.copy2(p, d / p.name)
    return d


def test_t1_discover_splices_returns_empty_under_min_pop(tmp_path):
    """With fewer than MIN_POP_SPLICE fibers, discover_splices returns []."""
    assert MIN_POP_SPLICE == 20, "MIN_POP_SPLICE drift — refresh this test"
    n = 12  # strictly < MIN_POP_SPLICE
    a_dir = _tiny_a_dir(tmp_path, n)
    fibers_a, _ = load_all(str(a_dir), "")
    assert 0 < len(fibers_a) <= n
    splices = discover_splices(fibers_a)
    assert splices == [], (
        "discover_splices should return [] when no km-bin clears "
        "MIN_POP_SPLICE; got %r" % splices
    )


def _qualifying_bend_fibers():
    """Synthesize fibers_a + fibers_b where ONE A event passes every
    gate inside scan_a_standalone_events up to line ~2968 (where the
    crash happens once splices==[]).

    The path: dist_km>=1, before EOL, before tailbox, loss>=BEND_THRESHOLD,
    B-side has matching event with positive loss → bidir_avg>=bt → reach
    the `target_sp = splices[best_si]` line where best_si is None.
    """
    bt = engine.BEND_THRESHOLD
    # Use a mid-span bend event at 50 km on a 100 km span.
    a_events = [
        {"dist_km": 0.0, "type": "0F0", "is_end": False, "splice_loss": 0.0,
         "is_reflective": False, "reflection": 0.0},
        {"dist_km": 50.0, "type": "0F1", "is_end": False,
         "splice_loss": max(bt + 0.05, 0.20),
         "is_reflective": False, "reflection": 0.0},
        {"dist_km": 100.0, "type": "1E9999", "is_end": True,
         "splice_loss": 0.0, "is_reflective": False, "reflection": 0.0},
    ]
    # B mirror: 100 - 50 = 50 km in B-frame
    b_events = [
        {"dist_km": 0.0, "type": "0F0", "is_end": False, "splice_loss": 0.0,
         "is_reflective": False, "reflection": 0.0},
        {"dist_km": 50.0, "type": "0F1", "is_end": False,
         "splice_loss": max(bt + 0.05, 0.20),
         "is_reflective": False, "reflection": 0.0},
        {"dist_km": 100.0, "type": "1E9999", "is_end": True,
         "splice_loss": 0.0, "is_reflective": False, "reflection": 0.0},
    ]
    rec_a = {"filename": "F0001.sor", "events": a_events, "_source": "sor"}
    rec_b = {"filename": "F0001.sor", "events": b_events, "_source": "sor"}
    return {1: rec_a}, {1: rec_b}, 100.0


@pytest.mark.xfail(
    strict=True,
    reason="engine pre-existing: empty splices crashes downstream passes "
           "(scan_a_standalone_events line ~2968 — `splices[best_si]` "
           "indexed with best_si=None when closure_centers is empty)",
)
def test_t1_scan_a_standalone_does_not_crash_on_empty_splices():
    """When discover_splices→[], the downstream pass should be a no-op,
    not raise TypeError.  XFAIL until the engine handles empty splices
    defensively.  When fixed → XPASS → strict=True fails the run →
    drop this marker."""
    fibers_a, fibers_b, total_span_a = _qualifying_bend_fibers()
    # Should NOT raise — but currently does.
    scan_a_standalone_events(
        fibers_a, splices=[], existing_results={},
        total_span_a=total_span_a, fibers_b=fibers_b,
    )


# ──────────────────────────────────────────────────────────────────────
# T2  filename sanitization for orig_loc / term_loc
# ──────────────────────────────────────────────────────────────────────
_FS_RESERVED = set('\\/:*?"<>|')


@pytest.mark.xfail(
    strict=True,
    reason="H4 pending: desktop_app builds f'splice_report_{a_name}_to_{b_name}.xlsx' "
           "from raw os.path.basename — reserved chars break download.  "
           "Fix should add _safe_filename_part() (or re.sub(r'[\\\\/:*?\"<>|]+', '_', s)) "
           "in splicereportmatchexfo and desktop_app should use it.",
)
def test_t2_filename_sanitize_strips_reserved_chars():
    """Once the engine exposes _safe_filename_part, no reserved fs char
    survives.  XFAIL until the symbol exists."""
    try:
        from splicereportmatchexfo import _safe_filename_part  # noqa: F401
    except ImportError as e:
        pytest.fail(f"_safe_filename_part not yet implemented: {e}")
    cases = [
        "Seattle/Stevens Pass",
        "Site*A",
        "A:Loc",
        'weird"name<>|here?',
    ]
    for raw in cases:
        out = _safe_filename_part(raw)
        bad = _FS_RESERVED & set(out)
        assert not bad, (
            f"sanitize({raw!r}) -> {out!r} still contains "
            f"reserved chars: {sorted(bad)}"
        )
        # Should be non-empty and stable on already-clean input
        assert out.strip(), f"sanitize({raw!r}) produced empty/whitespace"


# ──────────────────────────────────────────────────────────────────────
# T3  H2 — apply_field_gainer_rule runs before off-splice splitter,
#         but the splitter's eligible-event filter doesn't include
#         is_gainer.  Confirm the current (buggy) attribution and
#         XFAIL the desired behaviour.
# ──────────────────────────────────────────────────────────────────────
def _gainer_synthetic_setup():
    """Construct minimal all_results + splices with one off-splice gainer.

    Splice at 30 km, total span 100 km.  Event sits at 50 km (well
    beyond CLOSURE_MATCH_KM=150 m and outside launch/tailbox).
    """
    splices = [{
        "position_km": 30.0,
        "position_km_refined": 30.0,
        "column_kind": "splice",
    }]
    key = (5, 0)  # fiber 5, splice_idx 0
    all_results = {
        key: {
            "fiber": 5, "splice_idx": 0,
            "a_loss": 0.30, "b_loss": -0.32, "bidir_loss": -0.01,
            "bidir_dist": 50.0,
            "is_gainer": True, "is_flagged": True,
            "is_bend": False, "is_break": False, "is_broke": False,
            "is_bfill": False, "is_dead_zone": False,
            "is_a_only": False, "is_b_only": False, "is_ref": False,
            "event_type": "0F1", "event_source": "bidir",
            "label": "F5 -0.010 bidi (gainer)",
        }
    }
    return all_results, splices, key


def test_t3_h2_gainer_currently_stays_in_splice_column():
    """H2 BUG CONFIRMATION: today, an off-splice gainer keeps its
    original splice_idx because the splitter filter omits is_gainer.

    The day H2 gets fixed (add `or r.get('is_gainer')` to the splitter
    filter at line 1704-ish), this test will fail — replace it with the
    paired XFAIL below, which will then XPASS → flip both.
    """
    all_results, splices, key = _gainer_synthetic_setup()
    new_results, new_splices = split_offsplice_events_into_own_columns(
        all_results, splices, total_span_km=100.0,
    )
    # Bug present: only one splice column, gainer still under it.
    assert len(new_splices) == 1, (
        "Splitter created a phantom column for the gainer — looks like "
        "H2 was FIXED.  Update the test: drop this assertion and remove "
        "the @xfail from test_t3_h2_gainer_should_relocate."
    )
    assert key in new_results
    assert new_results[key].get("is_gainer") is True


@pytest.mark.xfail(
    strict=True,
    reason="H2 pending: split_offsplice_events_into_own_columns filter "
           "at ~line 1704 omits is_gainer, so off-splice gainers stay "
           "attributed to the wrong closure.  Fix: add "
           "`or r.get('is_gainer')` to the filter list.",
)
def test_t3_h2_gainer_should_relocate_to_own_column():
    """The behaviour we WANT: an off-splice gainer at 50 km lands in a
    phantom column at ~50 km, not in the splice at 30 km."""
    all_results, splices, _ = _gainer_synthetic_setup()
    new_results, new_splices = split_offsplice_events_into_own_columns(
        all_results, splices, total_span_km=100.0,
    )
    assert len(new_splices) >= 2, (
        "expected the gainer to spawn its own phantom column, "
        f"got splices={new_splices}"
    )
    # The gainer should now sit under a non-original column near 50 km.
    gainer_entries = [r for r in new_results.values() if r.get("is_gainer")]
    assert gainer_entries, "gainer was deleted entirely (unexpected)"
    g = gainer_entries[0]
    new_col = new_splices[g["splice_idx"]]
    new_col_km = new_col.get("position_km_refined", new_col["position_km"])
    assert abs(new_col_km - 50.0) < 0.5, (
        f"gainer relocated to wrong column @ {new_col_km} km"
    )


# ──────────────────────────────────────────────────────────────────────
# T4  H3 — Pass-2d B-side breaks must override Pass-2c B-fill
# ──────────────────────────────────────────────────────────────────────
def _bside_break_setup():
    """Build a minimal fibers_a / fibers_b for ONE fiber where:
        - A trace passes through (ends near total_span_a)
        - B trace breaks mid-span at b_eof = 30 km in B-frame
          → mirrors to A-frame 100 - 30 = 70 km
        - A splice exists at 70 km (closest column)
        - existing_results carries a Pass-2c is_bfill entry at that cell

    Pass-2d should overwrite the is_bfill entry with is_broke=True.
    """
    total_span_a = 100.0
    fnum = 7
    # A trace: a few events, ends at 100 km (intact A side)
    fibers_a = {
        fnum: {
            "filename": "F0007.sor",
            "events": [
                {"dist_km": 0.0, "type": "0F0", "is_end": False,
                 "splice_loss": 0.0},
                {"dist_km": 100.0, "type": "1E9999", "is_end": True,
                 "splice_loss": 0.0},
            ],
        },
        # Add a 2nd fiber so total_span_b median has data
        8: {
            "filename": "F0008.sor",
            "events": [
                {"dist_km": 0.0, "type": "0F0", "is_end": False,
                 "splice_loss": 0.0},
                {"dist_km": 100.0, "type": "1E9999", "is_end": True,
                 "splice_loss": 0.0},
            ],
        },
    }
    # B trace: breaks at 30 km (B-frame).  Mirrors to A-frame 70 km.
    fibers_b = {
        fnum: {
            "filename": "F0007.sor",
            "events": [
                {"dist_km": 0.0, "type": "0F0", "is_end": False,
                 "splice_loss": 0.0},
                {"dist_km": 30.0, "type": "1E9999", "is_end": True,
                 "splice_loss": 0.0},
            ],
        },
        # 2nd fiber: B intact, lets total_span_b median compute to 100.
        8: {
            "filename": "F0008.sor",
            "events": [
                {"dist_km": 0.0, "type": "0F0", "is_end": False,
                 "splice_loss": 0.0},
                {"dist_km": 100.0, "type": "1E9999", "is_end": True,
                 "splice_loss": 0.0},
            ],
        },
    }
    splices = [{
        "position_km": 70.0,
        "position_km_refined": 70.0,
        "column_kind": "splice",
    }]
    # Simulate Pass-2c having already written a B-fill entry at this cell.
    existing_results = {
        (fnum, 0): {
            "fiber": fnum, "splice_idx": 0,
            "bidir_loss": 0.30, "a_loss": None, "b_loss": 0.30,
            "bidir_dist": 70.0,
            "is_break": False, "is_broke": False, "is_bend": False,
            "is_bfill": True, "is_dead_zone": False,
            "is_a_only": False, "is_b_only": False,
            "is_flagged": True, "event_source": "bfill",
            "event_type": "0F1", "label": "F7 0.300 (B-fill)",
        },
    }
    return fibers_a, fibers_b, splices, existing_results, total_span_a


def test_t4_h3_pass2d_currently_skips_bfill_cell():
    """Current behaviour: scan_b_side_breaks refuses to overwrite the
    Pass-2c B-fill entry because the override condition only fires on
    is_dead_zone.  This is the H3 bug — confirm it, then XFAIL the
    desired behaviour below.
    """
    fa, fb, splices, existing, total_a = _bside_break_setup()
    new = scan_b_side_breaks(fa, fb, splices, existing, total_a)
    # Bug present: Pass-2d found nothing new at (7, 0).
    assert (7, 0) not in new, (
        "Pass-2d wrote at (7, 0) — looks like H3 was FIXED.  Update "
        "this test: drop this assertion + un-xfail the next one."
    )


@pytest.mark.xfail(
    strict=True,
    reason="H3 pending: scan_b_side_breaks line ~3403 — "
           "`if prior is not None and not prior.get('is_dead_zone'): continue` "
           "should also allow override when prior.get('is_bfill') (Pass-2d "
           "concrete broke is more informative than Pass-2c B-fill).",
)
def test_t4_h3_pass2d_should_override_bfill():
    """After fix: at (fiber, splice_idx) cell, Pass-2d's is_broke=True
    entry replaces Pass-2c's is_bfill entry."""
    fa, fb, splices, existing, total_a = _bside_break_setup()
    new = scan_b_side_breaks(fa, fb, splices, existing, total_a)
    assert (7, 0) in new, (
        f"Pass-2d should have written at (7, 0); got keys={list(new)}"
    )
    entry = new[(7, 0)]
    assert entry.get("is_broke") is True
    assert not entry.get("is_bfill", False)


# ──────────────────────────────────────────────────────────────────────
# T5  reburn-summary cosmetics on a fully-damaged span
# ──────────────────────────────────────────────────────────────────────
def test_t5_reburn_summary_zero_real_splices_renders_zero():
    """Current behaviour: when every column is bend/damage (no real
    splices), n_real_splices==0, total_cells==0, reburn_percentage==0.0."""
    from reburn_summary import compute_reburn_summary
    splices = [
        {"position_km": 10.0, "column_kind": "damage"},
        {"position_km": 30.0, "column_kind": "bend"},
    ]
    summary = compute_reburn_summary(
        all_results={}, splices=splices,
        n_fibers=144, ribbon_size=12,
    )
    assert summary["n_real_splices"] == 0
    assert summary["total_cells"] == 0
    assert summary["reburn_percentage"] == 0.0


@pytest.mark.xfail(
    strict=True,
    reason="cosmetic pending: when n_real_splices==0, the summary "
           "should expose an 'N/A — no real splice columns' marker "
           "instead of a misleading 0.00%.  Add a flag like "
           "summary['display_percentage'] = 'N/A — no real splice columns'.",
)
def test_t5_reburn_summary_zero_real_splices_renders_na():
    from reburn_summary import compute_reburn_summary
    summary = compute_reburn_summary(
        all_results={},
        splices=[{"position_km": 10.0, "column_kind": "damage"}],
        n_fibers=144, ribbon_size=12,
    )
    display = summary.get("display_percentage") or ""
    assert "N/A" in display and "real splice" in display.lower()


# ──────────────────────────────────────────────────────────────────────
# T6  mixed JSON+SOR in one dir issues a visible WARN
# ──────────────────────────────────────────────────────────────────────
def _mixed_dir(tmp_path: Path) -> Path:
    d = tmp_path / "mixed"
    d.mkdir()
    # Copy one .sor
    sors = sorted(p for p in FIXTURE_A_DIR.iterdir()
                  if p.name.lower().endswith(".sor"))
    shutil.copy2(sors[0], d / sors[0].name)
    # Drop a stub .json (parse failure is fine — we only test WARN emission
    # about MIXED-format presence, not parser success).
    (d / "F0099.json").write_text('{"not": "real otdr"}')
    return d


@pytest.mark.xfail(
    strict=True,
    reason="MED pending: _dir_has_json (engine line ~730) picks ONE "
           "format per dir and silently skips the rest.  Fix should "
           "print a WARN line naming the count of skipped files (e.g. "
           "'WARN: mixed JSON+SOR in <dir>; using JSON, skipped 1 SOR').",
)
def test_t6_mixed_json_sor_warns(tmp_path):
    d = _mixed_dir(tmp_path)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        load_all(str(d), str(FIXTURE_B_DIR))
    out = buf.getvalue().lower()
    # Required: a WARN line that specifically identifies the MIXED-format
    # condition (not just a generic per-file parse failure).  Today
    # _dir_has_json silently picks ONE format; once the fix lands, the
    # log will name the count + format of skipped files.
    assert "warn" in out and "mixed" in out and ("skipped" in out
                                                  or "ignored" in out), (
        "expected a visible WARN identifying the mixed-format skip; "
        f"stdout was:\n{buf.getvalue()}"
    )


# ──────────────────────────────────────────────────────────────────────
# T7  UTC midnight bucketing
# ──────────────────────────────────────────────────────────────────────
def _ts_record(filename: str, epoch_ts: int) -> dict:
    """Minimal fiber record shape audit_acquisition needs."""
    return {
        "filename": filename,
        "date_time": epoch_ts,
        "otdr_model": "FTBx",
        "otdr_serial": "SN-1",
        "_source": "sor",
        "events": [],
        "exfo_calibration": {"NominalPulseWidth": 100e-9,
                              "NumberOfAverages": 1024},
        "exfo_wavelength_nm": 1550.0,
        "duration_sec": 30.0,
    }


def test_t7_utc_midnight_currently_splits_pt_evening_into_two_days():
    """A single PT-evening shoot straddling UTC midnight currently
    shows as two calendar days in the audit verdict — a false warning
    to the tech.  Confirm the bug; XFAIL the desired behaviour below.
    """
    from acquisition_audit import audit_acquisition
    # 2026-06-12 23:30 UTC  and  2026-06-13 00:30 UTC
    # That's the same PT evening (~16:30 and 17:30 PT on 2026-06-12).
    fa = {1: _ts_record("F0001.sor", 1781306200),  # 2026-06-12 23:30 UTC
          2: _ts_record("F0002.sor", 1781309800)}  # 2026-06-13 00:30 UTC
    audit = audit_acquisition(fa, {})
    date_row = audit["file_fields"][0]
    assert date_row["name"].startswith("Test date")
    verdict = date_row["result"]
    # Bug: 2 distinct calendar days seen → verdict reports differ.
    consistent = verdict.get("consistent") if isinstance(verdict, dict) else None
    # Permissive on shape; ANY of these indicates "two days were
    # bucketed" — which is the bug we want to flag.
    txt = repr(verdict).lower()
    assert (consistent is False
            or "differ" in txt
            or "outlier" in txt
            or "2026-06-13" in txt and "2026-06-12" in txt), (
        f"expected current behaviour to flag two days; verdict={verdict!r}"
    )


@pytest.mark.xfail(
    strict=True,
    reason="MED pending: acquisition_audit._calendar_day buckets on UTC "
           "midnight, so a PT-evening shoot that crosses 00:00 UTC reads "
           "as two days.  Fix: bucket on the timezone of the OTDR "
           "(or apply a ±N-hour tolerance) so adjacent UTC days within "
           "a few hours register as consistent.",
)
def test_t7_utc_midnight_should_tolerate_local_day():
    from acquisition_audit import audit_acquisition
    fa = {1: _ts_record("F0001.sor", 1781306200),
          2: _ts_record("F0002.sor", 1781309800)}
    audit = audit_acquisition(fa, {})
    verdict = audit["file_fields"][0]["result"]
    consistent = verdict.get("consistent") if isinstance(verdict, dict) else None
    assert consistent is True, (
        f"expected verdict to mark as consistent after local-day "
        f"tolerance; got {verdict!r}"
    )


# ──────────────────────────────────────────────────────────────────────
# T8  engine-API signature contract
# ──────────────────────────────────────────────────────────────────────
def _param_names(fn) -> list[str]:
    return list(inspect.signature(fn).parameters.keys())


def test_t8_detect_launch_issues_signature():
    names = _param_names(detect_launch_issues)
    # Positional / kw arg names the UI passes by name.
    for required in ("fibers_a", "fibers_b", "first_splice_km",
                      "high_loss_db", "bad_refl_db"):
        assert required in names, (
            f"detect_launch_issues lost kw arg '{required}'; "
            f"current params={names}"
        )


def test_t8_build_ribbon_data_signature():
    sig = inspect.signature(build_ribbon_data)
    names = list(sig.parameters.keys())
    for required in ("results", "n_fibers", "ribbon_size",
                      "n_splices", "launch_issues"):
        assert required in names, (
            f"build_ribbon_data lost arg '{required}'; "
            f"current params={names}"
        )
    # launch_issues must be optional (UI sometimes passes None).
    assert sig.parameters["launch_issues"].default is None


def test_t8_write_xlsx_signature():
    sig = inspect.signature(write_xlsx)
    names = list(sig.parameters.keys())
    for required in ("cells", "splices", "n_fibers", "ribbon_size",
                      "output_path", "site_a", "site_b", "span_km",
                      "launch_cells_a", "launch_cells_b",
                      "fibers_a", "fibers_b", "all_results"):
        assert required in names, (
            f"write_xlsx lost arg '{required}'; current params={names}"
        )
