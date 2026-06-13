Test fixtures — tiny real OTDR span
====================================

Contents
--------
  span_A/  36 SOR files (LAGDUR0001.sor .. LAGDUR0036.sor)
           A-direction (La Grande → Durkee), fibers 1-36 (three ribbons).
  span_B/  36 SOR files (DURLAG0001.sor .. DURLAG0036.sor)
           B-direction (Durkee → La Grande), fibers 1-36 (three ribbons).

Total size: ~7.4 MB.

Why 36 fibers (not 12)?
-----------------------
The engine's ``discover_splices`` keeps a 1 km bin only when at least
``MIN_POP_SPLICE`` (=20, see splicereportmatchexfo.py) fibers land an
event in that bin.  A single ribbon (12 fibers) never clears the gate,
so the splice list comes back empty and downstream
``scan_a_standalone_events`` blows up indexing into it.  Three ribbons
(36 fibers) reliably yields ~4-5 closures and exercises the full
pipeline.

Source
------
Extracted from a real customer span:
  Original zip:  /Users/robertcolbert/Downloads/Span 7 La Grande to Durkee-selected/
                    {La Grande to Durkee, Durkee to La Grande}.zip
  Staging dir:   /tmp/durkee_run/{A,B}/...

We took three ribbons (fibers 1-36) of each direction.  The full span has 432
fibers per side, ~46 MB — way too big to vendor.  Three ribbons is the
minimum that lets the engine:
    * load both directions
    * discover closures (the engine's MIN_POP_SPLICE=20 gate eats single
      ribbons whole — see "Why 36 fibers" above)
    * detect launch issues (this is what regressed in commit e1d5692)
    * build the ribbon grid + write all four sheets

What it's used for
------------------
The end-to-end "Generate report" test (test_e2e_generate.py) runs the entire
desktop UI through Streamlit's AppTest harness against these fixtures and
verifies that:
    * the run finishes without raising
    * a workbook is written with the expected 4 sheets in the expected order
    * the Acquisition Parameters sheet is active (the workbook opens on it)
    * the Reburn Summary headline cells are populated
This catches signature drift between desktop_app.py and the engine — the
class of bug that shipped in commit e1d5692 (`detect_launch_issues` was
called with the wrong signature and unpacked the dict's keys into two
variables, raising `ValueError: too many values to unpack`).

Do not edit these SOR files.  If they need to be refreshed, re-copy from
/tmp/durkee_run or re-extract from the customer zip above.
