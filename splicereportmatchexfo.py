#!/usr/bin/env python3
"""
splicereportmatchexfo_StevensReport_withBends_April20.py
=========================================================

Splice QC report that:
  • MATCHES EXFO FastReporter's bidirectional analysis via JSON-trace wide-LSA
    (the algorithm that matches STEVEN'S REPORT on NEW-ELM 1152 within 0.001 dB
    for the six target fibers F110, F161, F217, F468, F1024, F133).
  • Adds BEND detection on top (ZeroDBIFTHEN Flag 3 rule).
  • Adds LAUNCH-ISSUE detection so fibers broken / damaged at the launch end
    can't silently disappear from the report.

APRIL 20 REVISION:
  - Built for the MIL-ELM data set Zach found bend losses in that the earlier
    splice report was misclassifying as splice events.
  - MIL-ELM is a new cable; STEVEN HAS NOT REVIEWED IT.  The Steven-report
    reference in this script's name is to the ALGORITHM LINEAGE only (the
    EXFO-match / Stevens-match logic we validated on NEW-ELM), not to any
    ground-truth comparison on MIL-ELM.

What BEND detection adds:

  * Closure centers are refined from the coarse 1 km bin to the MODE of
    fiber event positions in that bin.
  * For every flagged event, we measure the distance from the A-direction
    event position to the refined closure center.
  * If the offset is > 150 m AND the loss is >= 0.020 dB, the event is
    classified as a BEND (not a splice reburn).  This matches the boss's
    method from ZeroDBIFTHEN Flag 3.
  * Bends render in teal on the Excel report (three shades by severity:
    WATCH, REVIEW, HIGH) with a "BEND" label and the offset in metres.
  * Bends do NOT terminate analysis — the script keeps walking through
    later splices for the same fiber, same as it would past a clean splice.

PASS 1 (same as splice report):
  - For each fiber at each known splice closure position:
      Find A event → find matching B event → compute bidirectional loss
      Flag A-only events (no B match) if A loss >= threshold

PASS 2 (new):
  - For each fiber, scan ALL B-direction events above threshold
  - Convert each to A-frame coordinates
  - Skip any already caught in Pass 1
  - Match to nearest splice position (within 1.5 km)
  - If matching A event found: compute bidirectional loss → label A+B
  - If no A event: flag as B-only

CELL LABELS:
  325 .172        — standard A+B bidirectional splice (same as original report)
  325 .340 (B)    — B-direction only saw this event; A-direction had nothing
  325 .285 (A)    — A-direction only saw this event; no matching B entry

COLORS:
  Pink   — A+B bidirectional reburn (loss >= threshold)
  Red    — Break / Broke (both treated as physical damage; reflective
            break or non-reflective crush both render with the red fill)
  Blue   — B-fill (B-direction loss past a break, A-direction blind)
  Yellow — A-only (A saw it, B did not)
  Purple — B-only (B saw it, A did not)

USAGE
-----
    python splicereportmatchexfo.py A_DIR/ B_DIR/ --output report.xlsx

OPTIONS
    --output PATH    Output Excel file (default: splice_report_exfo.xlsx)
    --threshold dB   Flag threshold (default 0.150)
    --site-a NAME    A-end site name (default TUL)
    --site-b NAME    B-end site name (default BAR)
    --ribbon-size N  Fibers per ribbon (default 12)

REQUIREMENTS
    pip install numpy openpyxl
    sor_reader324802a.py must be in same directory.
"""

import os
import sys
import argparse
from collections import defaultdict

import numpy as np

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    print("ERROR: pip install openpyxl"); sys.exit(1)

from sor_reader324802a import parse_sor_full
# JSON-based grey-value measurement — matches EXFO's internal LSA calculation
# (see json_reader.py for the algorithm details)
from json_reader import (
    parse_otdr_json,
    measure_grey_loss_from_json,
    load_all_json,
)


# ═══════════════════════════════════════════════════════════════════════
#  DEFAULTS
# ═══════════════════════════════════════════════════════════════════════

REBURN_THRESHOLD = 0.150   # dB — flag anything at or above
NOMINAL_SPLICE   = 0.159   # dB expected per splice
RIBBON_SIZE      = 12      # fibers per ribbon
POSITION_TOL     = 1.5     # km tolerance for matching A↔B events
MIN_POP_SPLICE   = 20      # minimum fibers to define a splice position
END_REGION_KM    = 3.0     # last N km considered "end of fiber"
LAUNCH_FIBER_MAX = 3.0     # km — max distance for launch connector detection

# Wide-LSA grey-value measurement windows (matches EXFO's FastReporter
# behavior — see discussion in json_reader.py / previous experiments):
GREY_LSA_OUTER_M = 5000    # m — outer LSA window on each side of splice
GREY_LSA_INNER_M = 60      # m — inner dead zone on each side of splice

# ── BEND detection (from ZeroDBIFTHEN's Flag 3 rule / boss's method) ─────────
#
#   An OTDR event is a BEND (not a splice) if it meets BOTH:
#     1. Its bidir / single-direction loss magnitude >= BEND_THRESHOLD
#     2. Its position is more than CLOSURE_MATCH_KM from the nearest *real*
#        splice closure center (not the splice report's coarse 1 km bin,
#        but the MODE of fiber event positions inside that bin).
#
# Bends are flagged in their own category.  Analysis of downstream splices
# continues normally past a bend — unlike a break, a bend does not make the
# trace blind past it.
#
BEND_THRESHOLD        = 0.090   # dB — minimum loss to call an event a "bend"
CLOSURE_MATCH_KM      = 0.150   # km — tight window; farther → classify as bend
# Severity tiers intentionally collapsed: any bend ≥ BEND_THRESHOLD is
# rendered with the same yellow fill.  The old WATCH / REVIEW / HIGH
# tiers are retained as constants only for backward-compatibility with
# any downstream code that still references them.
BEND_HIGH_LOSS        = BEND_THRESHOLD
BEND_REVIEW_LOSS      = BEND_THRESHOLD
# Histogram bin for the mode-based closure-center refinement:
CLOSURE_MODE_BIN_M    = 25      # m — bin width for position-mode histogram
CLOSURE_MODE_WINDOW_M = 75      # m — window around mode peak for median refinement

# ── APRIL 23 revision: closure validation (drop phantom splice columns) ──
# A cluster discovered by discover_splices() is only a *real* closure if the
# fiber-event positions inside it cluster tightly.  Loose clusters are bend
# or damage zones mis-identified as splices.  Matches the tech's mental
# model: real closures look tight, bend zones are smeared.
CLOSURE_VALID_STD_MAX_M  = 150    # m — std of positions inside closure
CLOSURE_VALID_TIGHT_FRAC = 0.13   # fraction of ALL fibers whose event is
                                   #   within CLOSURE_MATCH_KM of the refined
                                   #   center.  Set intentionally low because
                                   #   broken fibers reduce the denominator:
                                   #   catastrophic breaks kill the count for
                                   #   every closure past the break zone.
                                   #   0.13 drops clear damage/bend phantoms
                                   #   (typical tight_frac ≈ 0.05–0.12) while
                                   #   keeping real splices (≥ 0.15).

# Additional April 23 closure-validity rules derived from raw-data analysis of
# the Cle Elum → Yakima 18.89 km bend zone (tech flagged bends; we flagged
# splice).  A real splice closure should look like random fusion-loss
# variation — small median loss, SOME apparent gainers from MFD mismatch.
# A bend zone looks very different: every fiber shows positive (loss) with
# a higher median.
CLOSURE_VALID_MIN_GAINER_FRAC = 0.05   # ≥ 5% of fibers in the cluster must
                                        #   show loss < 0 (apparent gain).
                                        #   Real splices ≈ 20-40% gainers.
                                        #   Bend zones ≈ 0% gainers.
CLOSURE_VALID_MEDIAN_LOSS_MAX = 0.100   # dB — median loss inside the tight
                                        #   cluster.  Real splices ≈ 0.04-
                                        #   0.08 dB median, bend zones push
                                        #   well above 0.10 dB.

# ── LAUNCH-issue detection (fibers broken/damaged at/near the launch end) ───
#
# Fibers with launch issues often silently disappear from the splice report:
# the event table ends almost immediately, so neither Pass 1 (splice analysis
# at known closures) nor Pass 2 (B-direction scan) has anything to match on.
# These fibers need to be flagged on their own so the tech knows to go look.
#
LAUNCH_HIGH_LOSS_DB          = -0.5   # signed launch-event loss threshold.  Rule:
                                      #   flag when launch_loss >= -0.5 dB (i.e.
                                      #   value is at -0.5 or anywhere closer to
                                      #   zero / positive).  Healthy MFD-mismatch
                                      #   launches show a strong gainer (more
                                      #   negative than -0.5 dB); anything weaker
                                      #   than that gainer signature is anomalous.
LAUNCH_BAD_REFL_DB           = -15.0  # launch reflectance stronger (closer to 0)
                                      #   than this → flag damaged/dirty connector.
                                      #   Refl is reported negative; healthy buried
                                      #   launch is -50 to -55 dB, damaged is
                                      #   -10 to -30, missing is < -70.  Rule
                                      #   (refl > -15) flags only the 0 to -15
                                      #   range — anomalously strong reflections
                                      #   that indicate connector damage.
# ── FIELD-EVENT GAINER GATE ─────────────────────────────────────────────────
# Mid-span events whose signed loss falls in the [-0.7, 0] dB range get
# flagged as suspicious gainers — these are weak-gainer / near-zero events
# that the regular bend rule would either miss (if |loss| < 0.090) or
# silently fold into a normal splice column.  Excludes anything inside
# LAUNCH_FIBER_MAX km of the launch or END_REGION_KM km of the fiber end —
# those zones are evaluated separately by the launch / end logic.
FIELD_GAINER_MIN_DB          = -0.7   # most-negative loss that still flags
FIELD_GAINER_MAX_DB          = 0.0    # least-negative loss that still flags
LAUNCH_REFL_OUTLIER_DB       = 10.0   # |fiber_refl − population_median| > this → issue
LAUNCH_NO_FIRST_SPLICE_TOL_KM = 2.0   # km — must see an event within this of the
                                      #      first population closure


# ═══════════════════════════════════════════════════════════════════════
#  AUTO-DETECT & NORMALIZE UNTRIMMED TRACES
# ═══════════════════════════════════════════════════════════════════════

def _normalize_untrimmed_events(events):
    """Detect and normalize events from SOR files where start/stop was not picked.

    Untrimmed pattern (tech did NOT pick start/stop):
      #1  1F  0.000 km   — OTDR port (instrument origin)
      #2  1F  ~1.0  km   — launch connector (fiber-under-test starts here)
      ...  splice events with ~1 km offset  ...
      #N-1  1F  ~98.3 km — far-end connector (receive fiber)
      #N    xE  ~99.3 km — end-of-fiber marker

    Trimmed pattern (tech already picked start/stop):
      #1  1F  0.000 km   — launch connector (already set as origin)
      ...  splice events at correct positions  ...
      #N    xE  ~97.2 km — end-of-fiber marker

    Detection: first TWO events are both reflective (1F) non-end events,
    with the second one at a short distance (< LAUNCH_FIBER_MAX km).

    Normalization:
      1. Remove event #1 (OTDR port)
      2. Re-reference all distances from the launch connector (event #2 → dist 0)
      3. Remove the far-end connector (last 1F within 3 km of the end event)
    """
    if len(events) < 3:
        return events

    # ── Detect untrimmed ──
    e0, e1 = events[0], events[1]
    if not (e0['is_reflective'] and not e0['is_end'] and
            e0['time_of_travel'] == 0 and
            e1['is_reflective'] and not e1['is_end'] and
            0 < e1['dist_km'] < LAUNCH_FIBER_MAX):
        return events  # already trimmed — no-op

    launch_dist = e1['dist_km']
    launch_travel = e1['time_of_travel']

    # ── Find end-of-fiber event ──
    end_idx = None
    for i, e in enumerate(events):
        if e['is_end']:
            end_idx = i
            break

    # ── Find far-end connector: last 1F just before the end event ──
    far_end_idx = None
    if end_idx is not None and end_idx > 1:
        end_dist = events[end_idx]['dist_km']
        for i in range(end_idx - 1, 0, -1):
            if events[i]['is_reflective'] and not events[i]['is_end']:
                if (end_dist - events[i]['dist_km']) < LAUNCH_FIBER_MAX:
                    far_end_idx = i
                break  # only check the immediately preceding reflective event

    # ── Compute the adjusted end position ──
    # When the tech picks start/stop, the end is set at the far-end connector,
    # not at the trace noise floor beyond the receive fiber.  Mirror that by
    # moving the end event to the far-end connector position.
    far_end_norm_dist = None
    far_end_norm_travel = None
    if far_end_idx is not None:
        far_end_norm_dist = round(events[far_end_idx]['dist_km'] - launch_dist, 4)
        far_end_norm_travel = max(0, events[far_end_idx]['time_of_travel'] - launch_travel)

    # ── Build normalized event list ──
    normalized = []
    for i, e in enumerate(events):
        if i == 0:           # skip OTDR port
            continue
        if i == far_end_idx: # skip far-end connector
            continue
        new_e = dict(e)
        new_e['dist_km'] = round(e['dist_km'] - launch_dist, 4)
        new_e['time_of_travel'] = max(0, e['time_of_travel'] - launch_travel)
        # Move end event to the far-end connector position (strip receive fiber)
        if e['is_end'] and far_end_norm_dist is not None:
            new_e['dist_km'] = far_end_norm_dist
            new_e['time_of_travel'] = far_end_norm_travel
        normalized.append(new_e)

    return normalized


# ═══════════════════════════════════════════════════════════════════════
#  TRACE-BASED SPAN & BREAK DETECTION
# ═══════════════════════════════════════════════════════════════════════

# Detection thresholds
SPIKE_MIN_DB      = 3.0    # minimum dB above baseline for connector spike
NOISE_STDDEV_THR  = 0.5    # stddev threshold separating signal from noise floor
BREAK_MEAN_THR    = 58.0   # dB mean threshold for break/saturation
BREAK_STDDEV_THR  = 1.0    # stddev threshold for noise spike at break
NOISE_WINDOW      = 50     # samples for sliding window statistics


def _sample_to_km(idx, ior, pts, acq_range):
    """Convert a trace sample index to distance in km."""
    return idx * 0.02998 * 2 * acq_range / (1000.0 * ior * pts)


def _km_to_sample(km, ior, pts, acq_range):
    """Convert distance in km to a trace sample index."""
    return int(round(km * 1000.0 * ior * pts / (0.02998 * 2 * acq_range)))


def _sliding_stats(trace, window=NOISE_WINDOW):
    """Compute sliding-window mean and stddev for the trace.

    Returns (means, stds) arrays of same length as trace.
    Uses a fast cumulative-sum approach.
    """
    n = len(trace)
    means = np.empty(n)
    stds = np.empty(n)
    for i in range(n):
        lo = max(0, i - window)
        hi = min(n, i + window + 1)
        seg = trace[lo:hi]
        means[i] = seg.mean()
        stds[i] = seg.std()
    return means, stds


def _detect_launch_from_trace(trace, pts, acq_range, ior):
    """Detect the launch connector from the raw OTDR trace.

    The launch connector creates a dead-zone DIP in the backscatter trace:
    the strong Fresnel reflection saturates the detector, causing a local
    minimum at ~0.5-2.0 km.  The fiber-under-test begins after this dip.

    Returns the sample index of the launch connector.
    """
    n = len(trace)
    # Skip the OTDR port dead zone (first ~0.1km) and search 0.1-3.0 km
    skip = _km_to_sample(0.1, ior, pts, acq_range)
    search_end = min(_km_to_sample(LAUNCH_FIBER_MAX, ior, pts, acq_range), n - 1)

    if search_end <= skip + 10:
        return 0

    region = trace[skip:search_end]
    # The launch connector is at the local minimum (bottom of the dead-zone dip)
    min_rel = int(np.argmin(region))
    launch_idx = skip + min_rel

    return launch_idx


def _detect_noise_floor_from_trace(trace, launch_idx, pts, acq_range, ior):
    """Find where the trace transitions from signal to noise floor.

    The OTDR pulse width limits how far the saved trace has clean signal.
    Beyond this point, the trace goes to saturation (~64 dB) with high noise.

    This is the TRACE noise floor — NOT necessarily the fiber end (the OTDR
    events extend further due to multi-acquisition).  Used for break detection:
    if a fiber's trace goes to noise significantly earlier than the population,
    it likely has a break.

    Returns the sample index where signal transitions to noise.
    """
    n = len(trace)
    step = max(1, NOISE_WINDOW // 2)

    # Scan backward from end to find where clean signal begins
    for i in range(n - NOISE_WINDOW - 1, launch_idx, -step):
        seg = trace[i:i + NOISE_WINDOW]
        if seg.std() < NOISE_STDDEV_THR and seg.mean() < BREAK_MEAN_THR:
            # Found clean signal — noise floor starts after this
            noise_start = i + NOISE_WINDOW
            # Refine forward
            for j in range(noise_start, min(noise_start + NOISE_WINDOW * 2, n)):
                seg2 = trace[max(0, j - 10):j + 10]
                if seg2.std() > NOISE_STDDEV_THR or seg2.mean() > BREAK_MEAN_THR:
                    return j
            return noise_start

    return n - 1  # entire trace is noise (shouldn't happen)


def _detect_breaks_from_trace(trace, launch_idx, end_idx):
    """Detect mid-span breaks from the raw trace.

    A break is a sudden jump to near-saturation (>58 dB) with elevated noise,
    preceded by normal signal (<55 dB, low noise).

    Returns list of sample indices where breaks occur.
    """
    breaks = []
    n = len(trace)
    w = NOISE_WINDOW
    i = launch_idx + w

    while i < end_idx - w:
        seg_after = trace[i:i + w]
        seg_before = trace[max(launch_idx, i - w):i]

        mean_after = seg_after.mean()
        mean_before = seg_before.mean()
        std_after = seg_after.std()

        # Break: signal was present before, saturated after
        if (mean_before < 55.0 and
                mean_after > BREAK_MEAN_THR and
                std_after > BREAK_STDDEV_THR):
            # Walk backward to find the exact transition sample
            break_idx = i
            for j in range(i, max(launch_idx, i - w), -1):
                if trace[j] < 55.0:
                    break_idx = j + 1
                    break
            breaks.append(break_idx)
            # Skip past this break region
            i += w * 4
            continue

        i += w // 2

    return breaks


def _enhance_events_with_trace(fiber_result, expected_span_km, ior=None, pop_noise_floor_km=None):
    """Enhance a fiber's event list using raw trace analysis.

    Detects the launch connector from the trace (more accurate than events),
    detects breaks where the trace goes to saturation earlier than expected,
    and re-normalizes events using the trace-detected offset.

    For end-of-fiber: uses the EVENTS (not the trace) because the OTDR's
    multi-acquisition events extend further than the single-acquisition
    saved trace.

    Modifies fiber_result['events'] in place.
    """
    trace = fiber_result.get('full_trace')
    if trace is None:
        return
    pts = fiber_result['full_points']
    acq = fiber_result['acq_range']
    if ior is None:
        ior = fiber_result.get('ior', 1.4682)

    events = fiber_result['events']

    # ── Detect if this is an untrimmed file ──
    is_untrimmed = (len(events) >= 2 and
                    events[0]['is_reflective'] and not events[0]['is_end'] and
                    events[0]['time_of_travel'] == 0 and
                    events[1]['is_reflective'] and not events[1]['is_end'] and
                    0 < events[1]['dist_km'] < LAUNCH_FIBER_MAX)

    if not is_untrimmed:
        return  # already trimmed — no trace enhancement needed

    # ── Trace-based launch detection ──
    launch_idx = _detect_launch_from_trace(trace, pts, acq, ior)
    launch_km = _sample_to_km(launch_idx, ior, pts, acq)

    # ── Trace-based noise floor detection (for break detection) ──
    noise_floor_idx = _detect_noise_floor_from_trace(trace, launch_idx, pts, acq, ior)
    noise_floor_km = _sample_to_km(noise_floor_idx, ior, pts, acq)

    # ── Break detection from trace ──
    break_indices = _detect_breaks_from_trace(trace, launch_idx, noise_floor_idx)
    break_kms = [_sample_to_km(bi, ior, pts, acq) for bi in break_indices]

    # ── Compare this fiber's noise floor to the POPULATION noise floor ──
    # Normal fibers all hit noise at roughly the same distance (pulse width
    # limit).  A fiber whose trace goes to noise significantly earlier than
    # the population has a break/broke.
    trace_span = noise_floor_km - launch_km
    ref_noise_floor = pop_noise_floor_km if pop_noise_floor_km else expected_span_km
    is_trace_broke = (ref_noise_floor > 0 and
                      trace_span < ref_noise_floor - END_REGION_KM)

    # If trace indicates broke but break detector didn't find a specific break,
    # inject a break at the noise floor transition
    if is_trace_broke and not break_kms:
        break_kms.append(noise_floor_km)

    # ── End-of-fiber from EVENTS (not trace) ──
    # The events come from multi-acquisition and extend further than the saved trace.
    # Find end event and far-end connector from the event list.
    end_evt_idx = None
    for i, e in enumerate(events):
        if e['is_end']:
            end_evt_idx = i
            break

    # Find far-end connector: last 1F before end, close to end
    far_end_evt_idx = None
    if end_evt_idx is not None:
        end_dist = events[end_evt_idx]['dist_km']
        for i in range(end_evt_idx - 1, 0, -1):
            if events[i]['is_reflective'] and not events[i]['is_end']:
                if (end_dist - events[i]['dist_km']) < LAUNCH_FIBER_MAX:
                    far_end_evt_idx = i
                break

    # Fiber end = far-end connector position (or end event if no connector found)
    if far_end_evt_idx is not None:
        fiber_end_km = events[far_end_evt_idx]['dist_km']
    elif end_evt_idx is not None:
        fiber_end_km = events[end_evt_idx]['dist_km']
    else:
        fiber_end_km = noise_floor_km

    # ── Re-normalize events ──
    launch_travel = int(round(launch_idx * 2 * acq / pts))

    normalized = []
    for i, e in enumerate(events):
        if i == 0 and e['time_of_travel'] == 0:
            continue  # skip OTDR port
        if i == far_end_evt_idx:
            continue  # skip far-end connector
        new_e = dict(e)
        new_e['dist_km'] = round(e['dist_km'] - launch_km, 4)
        new_e['time_of_travel'] = max(0, e['time_of_travel'] - launch_travel)
        # Adjust end event to fiber end (far-end connector position)
        if e['is_end']:
            new_e['dist_km'] = round(fiber_end_km - launch_km, 4)
        normalized.append(new_e)

    # ── Inject synthetic break events from trace ──
    for bk_km in break_kms:
        bk_norm = round(bk_km - launch_km, 4)
        if bk_norm < 1.0:
            continue
        # Don't inject if there's already an end event before this position
        existing_end = [ne for ne in normalized if ne['is_end'] and ne['dist_km'] < bk_norm]
        if existing_end:
            continue

        # Add a break event (1F reflective with weak Fresnel)
        normalized.append({
            'number': 999,
            'time_of_travel': int(round((bk_km * 1000.0 * ior / 0.02998) * 2)),
            'dist_km': bk_norm,
            'splice_loss': 0.0,
            'reflection': -35.0,
            'slope': 0.0,
            'type': '1F9999LS',
            'is_reflective': True,
            'is_end': False,
        })
        # Remove any end events that are AFTER this break
        normalized = [ne for ne in normalized if not (ne['is_end'] and ne['dist_km'] > bk_norm)]
        # Add end event just after the break
        normalized.append({
            'number': 1000,
            'time_of_travel': int(round(((bk_km + 0.1) * 1000.0 * ior / 0.02998) * 2)),
            'dist_km': round(bk_norm + 0.1, 4),
            'splice_loss': 0.0,
            'reflection': 0.0,
            'slope': 0.0,
            'type': '0E9999LS',
            'is_reflective': False,
            'is_end': True,
        })

    # Sort by distance
    normalized.sort(key=lambda e: (e['dist_km'], 0 if not e['is_end'] else 1))

    fiber_result['events'] = normalized
    fiber_result['_trace_launch_km'] = launch_km
    fiber_result['_trace_end_km'] = fiber_end_km
    fiber_result['_trace_noise_floor_km'] = noise_floor_km
    fiber_result['_trace_breaks'] = break_kms
    fiber_result['_trace_is_broke'] = is_trace_broke


# ═══════════════════════════════════════════════════════════════════════
#  STEP 1 — Load all fibers
# ═══════════════════════════════════════════════════════════════════════

def _extract_fiber_num(fn):
    """Extract fiber number from a SOR/JSON filename (digits only)."""
    base = fn.split('.')[0].split(' ')[0]  # strip extension and trailing space
    base = base.split('_')[0]
    digits = ''.join(c for c in base if c.isdigit())
    return int(digits) if digits else None


def _dir_has_json(d):
    """True if directory contains any .json files."""
    if not d or not os.path.isdir(d):
        return False
    for fn in os.listdir(d):
        if fn.lower().endswith('.json'):
            return True
    return False


def load_all(dir_a, dir_b):
    """Load fibers from A and B directories.  Each directory can contain
    either SOR files or EXFO JSON exports — auto-detected per directory.
    When JSON is available it is preferred (it carries the same trace
    samples as SOR plus per-event LSA markers, per-section attenuation,
    and a cleaner event list for grey-value measurement)."""
    fibers_a, fibers_b = {}, {}

    def _load_dir(d, out):
        if not d or not os.path.isdir(d):
            return
        use_json = _dir_has_json(d)
        ext = '.json' if use_json else '.sor'
        parser = parse_otdr_json if use_json else (lambda p: parse_sor_full(p, trim=False))
        for fn in sorted(os.listdir(d)):
            if not fn.lower().endswith(ext):
                continue
            try:
                r = parser(os.path.join(d, fn))
            except Exception as exc:
                print(f"  WARN: failed to parse {fn}: {exc}")
                continue
            if not r:
                continue
            fnum = _extract_fiber_num(fn)
            if fnum:
                r['_source'] = 'json' if use_json else 'sor'
                out[fnum] = r

    _load_dir(dir_a, fibers_a)
    _load_dir(dir_b, fibers_b)
    return fibers_a, fibers_b


# ═══════════════════════════════════════════════════════════════════════
#  Helper: measure grey-value splice loss from a direction's JSON trace
# ═══════════════════════════════════════════════════════════════════════

def _grey_loss(fiber_data, splice_km):
    """Return the LSA-measured splice loss at `splice_km` from this fiber's
    trace, or None if not available (SOR-only data, or trace region bad).

    Uses wide-LSA (±5km outer, ±60m inner) matching EXFO's approach.
    Only applicable when the fiber was loaded from JSON (the SOR's
    single-acquisition trace typically saturates past ~34 km and can't
    support LSA beyond that distance)."""
    if fiber_data is None:
        return None
    if fiber_data.get('_source') != 'json':
        return None
    return measure_grey_loss_from_json(
        fiber_data, splice_km,
        outer_m=GREY_LSA_OUTER_M,
        inner_m=GREY_LSA_INNER_M,
    )


# ═══════════════════════════════════════════════════════════════════════
#  STEP 2 — Discover splice closure positions from the A-direction population
# ═══════════════════════════════════════════════════════════════════════

def discover_splices(fibers_a):
    bins = defaultdict(list)
    for fnum, r in fibers_a.items():
        for e in r['events']:
            if e['dist_km'] < 1.0 or e['is_end']: continue
            if not e['type'].startswith('0F') and not e['type'].startswith('1F'): continue
            bk = round(e['dist_km'])
            bins[bk].append(e['dist_km'])

    splices = []
    for bk in sorted(bins.keys()):
        if len(bins[bk]) < MIN_POP_SPLICE: continue
        avg_pos = round(np.mean(bins[bk]), 2)
        splices.append({'bin': bk, 'position_km': avg_pos, 'count': len(bins[bk])})

    # Merge bins within 1 km of each other
    merged = []
    for sp in splices:
        if merged and abs(sp['position_km'] - merged[-1]['position_km']) < 1.0:
            if sp['count'] > merged[-1]['count']:
                merged[-1] = sp
        else:
            merged.append(sp)

    return merged


# ═══════════════════════════════════════════════════════════════════════
#  STEP 2b — Refine closure centers using the MODE of fiber event positions
#            (lets us cleanly distinguish splices from bends)
# ═══════════════════════════════════════════════════════════════════════

def _classify_phantom(sp, fibers_a):
    """Classify a dropped phantom closure as 'bend' or 'damage' based on
    fiber behavior at its position.  Damage zones show many fibers that
    terminate near this km; bend zones have surviving fibers with large
    positive losses and no apparent gainers."""
    pos = sp.get('position_km_refined', sp['position_km'])
    n_broke_near = 0
    for r in fibers_a.values():
        end = [e for e in r['events'] if e.get('is_end')]
        if not end: continue
        eof = end[0]['dist_km']
        # Fiber terminates within ±500 m of this position → counts as damage
        if abs(eof - pos) < 0.5:
            n_broke_near += 1
    sp['broke_near_count'] = n_broke_near
    if n_broke_near >= 10:
        return 'damage'
    # Otherwise it's a bend zone (or near-empty)
    return 'bend'


def refine_closure_centers(fibers_a, splices, validate=True,
                           valid_std_max_m=None, valid_tight_frac=None,
                           valid_min_gainer_frac=None,
                           valid_median_loss_max=None,
                           return_phantoms=False):
    """Refine each splice center to the MODE of fiber events in a ±1 km
    window; optionally VALIDATE the cluster and drop phantom closures.

    Adds fields to each splice dict:
        position_km_refined : mode-based closure center (km)
        position_spread_m   : max − min of fiber event positions in window (m)
        position_std_m      : stddev of those positions (m)
        tight_frac          : fraction of fibers with an event within
                              CLOSURE_MATCH_KM of the refined center
        is_real_closure     : True iff tight enough to be a real splice

    When validate=True, clusters that fail the tightness tests are
    REMOVED from the returned list (April 23 revision: phantom closures
    at bend/damage zones no longer get splice columns).
    """
    std_max   = CLOSURE_VALID_STD_MAX_M      if valid_std_max_m      is None else valid_std_max_m
    tight_fr  = CLOSURE_VALID_TIGHT_FRAC     if valid_tight_frac     is None else valid_tight_frac
    min_gnr   = CLOSURE_VALID_MIN_GAINER_FRAC if valid_min_gainer_frac is None else valid_min_gainer_frac
    med_max   = CLOSURE_VALID_MEDIAN_LOSS_MAX if valid_median_loss_max is None else valid_median_loss_max
    n_fibers_total = len(fibers_a) or 1

    out = []
    dropped = []
    for sp in splices:
        center_guess = sp['position_km']
        nearby = []          # positions (km)
        nearby_losses = []   # losses (signed, dB) paired with nearby[]
        for r in fibers_a.values():
            for e in r['events']:
                if e['dist_km'] < 1.0 or e['is_end']:
                    continue
                if abs(e['dist_km'] - center_guess) < 1.0:
                    nearby.append(e['dist_km'])
                    nearby_losses.append(e.get('splice_loss') or 0.0)

        if not nearby:
            sp['position_km_refined'] = center_guess
            sp['position_spread_m'] = 0.0
            sp['position_std_m'] = 0.0
            sp['tight_frac'] = 0.0
            sp['is_real_closure'] = False
            if not validate:
                out.append(sp)
            else:
                dropped.append(sp)
            continue

        arr = np.array(nearby)
        # Mode peak
        bin_km = CLOSURE_MODE_BIN_M / 1000.0
        nbins = max(5, int(np.ceil((arr.max() - arr.min()) / bin_km)))
        hist, edges = np.histogram(arr, bins=nbins, range=(arr.min(), arr.max()))
        peak_idx = int(np.argmax(hist))
        peak_center = (edges[peak_idx] + edges[peak_idx + 1]) / 2.0

        local_mask = np.abs(arr - peak_center) < (CLOSURE_MODE_WINDOW_M / 1000.0)
        if local_mask.sum() >= 5:
            refined = float(np.median(arr[local_mask]))
        else:
            refined = float(peak_center)
        sp['position_km_refined'] = refined
        sp['position_spread_m']   = float(arr.max() - arr.min()) * 1000
        sp['position_std_m']      = float(np.std(arr)) * 1000

        # Tightness: fraction of fibers whose event is within ±CLOSURE_MATCH_KM
        tight_mask = np.abs(arr - refined) < CLOSURE_MATCH_KM
        tight_count = int(tight_mask.sum())
        sp['tight_frac'] = tight_count / float(n_fibers_total)
        # Use std within the tight zone for the quality check
        tight_std_m = float(np.std(arr[tight_mask])) * 1000 if tight_count > 3 else 999.0
        sp['tight_std_m'] = tight_std_m

        # Loss-distribution rules (April 23 anti-phantom-closure):
        # collect losses of events INSIDE the tight window; compute gainer
        # fraction and median loss
        loss_arr = np.array(nearby_losses)
        tight_losses = loss_arr[tight_mask]
        if len(tight_losses) >= 5:
            sp['gainer_frac']      = float((tight_losses < 0).sum() / len(tight_losses))
            sp['median_loss_db']   = float(np.median(tight_losses))
            sp['mean_loss_db']     = float(np.mean(tight_losses))
            sp['abs_median_loss']  = float(np.median(np.abs(tight_losses)))
        else:
            sp['gainer_frac']     = 0.0
            sp['median_loss_db']  = 0.0
            sp['mean_loss_db']    = 0.0
            sp['abs_median_loss'] = 0.0

        # Verdict: a closure is REAL unless it fails one of these tests.
        # The tight-std and tight-frac tests are independent fatal signals.
        # The loss-distribution tests (no-gainers + high-median) are COMBINED
        # into a single test — both must fail together for a phantom flag,
        # because a real splice between matched fiber lots could legitimately
        # show zero gainers, and by itself a slightly-elevated median is not
        # enough to drop a closure.  Requiring BOTH to fail at the same time
        # catches bend/damage zones (which always fail both) without risking
        # a real closure where the fibers happen to produce uniform positive
        # losses.
        # Validation now relies solely on the loss-distribution physics gate
        # (zero gainers AND elevated median = bend/damage signature).  The
        # older tight_std_m and tight_frac geometry gates were removed per
        # tech direction — they were dropping legitimate closures on long
        # cables where event scatter or low-participation clusters are
        # normal.  tight_std_m / tight_frac are still computed for diagnostic
        # display only.
        fails = []
        # Apply the loss-distribution test whenever the cluster is at least
        # as big as the MIN_POP_SPLICE candidate threshold (20 fibers).  The
        # older 50-fiber guard was too strict now that this is the only gate
        # — damage zones on long cables often have < 50 fibers in the tight
        # cluster because most of the population is already broken upstream.
        if len(tight_losses) >= MIN_POP_SPLICE:
            no_gainers_fail = sp['gainer_frac'] < min_gnr
            high_median_fail = sp['median_loss_db'] > med_max
            if no_gainers_fail and high_median_fail:
                fails.append(
                    f'loss_distribution(gainers={sp["gainer_frac"]:.2f} + '
                    f'median={sp["median_loss_db"]:+.3f}dB)'
                )
        sp['validation_fails'] = fails
        sp['is_real_closure'] = not fails

        if not validate or sp['is_real_closure']:
            out.append(sp)
        else:
            dropped.append(sp)

    if validate and dropped:
        print(f"  Dropped {len(dropped)} phantom closure(s) (bend/damage zones):")
        for sp in dropped:
            sp['phantom_type'] = _classify_phantom(sp, fibers_a)
            sp['column_kind'] = sp['phantom_type']    # 'bend' or 'damage'
            fail_str = ' + '.join(sp.get('validation_fails', [])) or 'no_data'
            print(f"    {sp['position_km']:8.2f} km  "
                  f"[{sp['phantom_type']}]  "
                  f"(tight_frac {sp['tight_frac']:.2f}, "
                  f"gainer_frac {sp['gainer_frac']:.2f}, "
                  f"median_loss {sp['median_loss_db']:+.3f} dB, "
                  f"broke_near {sp.get('broke_near_count',0)})  "
                  f"→ FAIL: {fail_str}")

    # Every kept closure is tagged as 'splice' for downstream column rendering
    for sp in out:
        sp['column_kind'] = 'splice'

    if return_phantoms:
        return out, dropped
    return out


def _is_bend_event(event_pos_km, splice_center_km, loss):
    """Apply ZeroDBIFTHEN Flag-3 rule: an event is a BEND if its loss is
    above BEND_THRESHOLD AND its position is more than CLOSURE_MATCH_KM
    away from the splice closure center."""
    if abs(loss) < BEND_THRESHOLD:
        return False
    return abs(event_pos_km - splice_center_km) > CLOSURE_MATCH_KM


def _is_field_gainer(event_pos_km, total_span_km, loss):
    """Field-gainer rule: an event is flagged as a field gainer if its
    signed loss falls in [FIELD_GAINER_MIN_DB, FIELD_GAINER_MAX_DB] (i.e.
    -0.7 to 0 dB by default) AND its position is mid-span — at least
    LAUNCH_FIBER_MAX km from the launch (so the launch-loss rule owns
    it) and at least END_REGION_KM km before the fiber end (so the
    end-of-fiber region doesn't pollute the flag)."""
    if not (FIELD_GAINER_MIN_DB <= loss <= FIELD_GAINER_MAX_DB):
        return False
    if event_pos_km < LAUNCH_FIBER_MAX:
        return False
    if total_span_km > 0 and event_pos_km > (total_span_km - END_REGION_KM):
        return False
    return True


def apply_field_gainer_rule(all_results, total_span_km):
    """Post-pass annotator.  For every result that carries a usable loss
    value and a usable position, check whether it lands in the field-
    gainer range; if so, flag it with is_gainer=True and clear is_bend
    (gainer takes priority over the geometric bend tag in the [-0.7, 0]
    overlap range, per tech direction).

    Skips results that are already break / broke / dead_zone — those
    classifications win unconditionally."""
    flagged = 0
    for key, r in all_results.items():
        if not isinstance(r, dict):
            continue
        if r.get('is_break') or r.get('is_broke') or r.get('is_dead_zone'):
            continue
        # Pick the most representative signed loss value.  Prefer the A
        # direction's signed loss; fall back to bidir / B as needed.
        a_loss = r.get('a_loss')
        b_loss = r.get('b_loss')
        bidir  = r.get('bidir_loss')
        loss_signed = None
        if a_loss is not None:
            loss_signed = a_loss
        elif bidir is not None:
            loss_signed = bidir
        elif b_loss is not None:
            loss_signed = b_loss
        if loss_signed is None:
            continue
        pos_km = r.get('bidir_dist') or r.get('a_dist') or r.get('position_km')
        if pos_km is None:
            continue
        if _is_field_gainer(pos_km, total_span_km, loss_signed):
            r['is_gainer'] = True
            # Gainer wins over bend in the overlap range — strip the
            # bend tag so the cell renders cleanly with the gainer fill.
            if r.get('is_bend'):
                r['is_bend'] = False
                # event_source still 'a_only' / 'bidir' / etc.
            r['is_flagged'] = True
            flagged += 1
    return flagged


def _bend_severity(loss):
    # Severity tiers collapsed — every bend ≥ BEND_THRESHOLD is simply 'BEND'.
    return 'BEND'


def _format_loss(val):
    """'.172' style — drops leading 0. like Steven's report."""
    s = f"{abs(val):.3f}"
    return s[1:] if s.startswith('0.') else s


# ═══════════════════════════════════════════════════════════════════════
#  STEP 2c — Detect launch-end issues (fibers broken / damaged at launch)
#
#  These fibers would otherwise be silent in the report because their
#  event tables are truncated immediately after the launch connector.
# ═══════════════════════════════════════════════════════════════════════

def _fiber_launch_info(r):
    """Extract launch-connector event info from a fiber's events.
    Returns (first_launch_event_dict | None, end_km | None, n_events)."""
    if r is None:
        return None, None, 0
    events = r.get('events') or []
    launch_evt = None
    if events and events[0].get('is_reflective') and events[0]['dist_km'] < 0.5:
        launch_evt = events[0]
    end_events = [e for e in events if e.get('is_end')]
    end_km = end_events[0]['dist_km'] if end_events else None
    return launch_evt, end_km, len(events)


def detect_launch_issues(fibers_a, fibers_b, first_splice_km=None,
                          high_loss_db=None, bad_refl_db=None,
                          **_ignored):
    """Return {fiber_num: launch_issue_dict} for every fiber that has a
    launch-end problem in either direction.

    Optional overrides (used by the Streamlit sidebar):
      high_loss_db  — launch-connector loss >= this flags HIGH_LAUNCH_LOSS
      bad_refl_db   — launch reflectance > this flags BAD_LAUNCH_REFL
    Any other kwargs are accepted and ignored for forward-compat.

    launch_issue_dict has:
      a_tags : list[str]   — issue tags for A direction (empty if none)
      b_tags : list[str]   — issue tags for B direction
      severity : 'HIGH' | 'REVIEW' | 'WATCH'
      summary : str        — human-readable label for the cell
    """
    hi_loss = LAUNCH_HIGH_LOSS_DB if high_loss_db is None else float(high_loss_db)
    bad_refl = LAUNCH_BAD_REFL_DB if bad_refl_db is None else float(bad_refl_db)
    # Population medians
    def _gather_launch_refls(fibers):
        refls = []
        for r in fibers.values():
            le, _, _ = _fiber_launch_info(r)
            if le is not None and le.get('reflection') is not None and le['reflection'] < 0:
                refls.append(le['reflection'])
        return float(np.median(refls)) if refls else None

    a_refl_median = _gather_launch_refls(fibers_a)
    b_refl_median = _gather_launch_refls(fibers_b)

    all_fibers = sorted(set(fibers_a.keys()) | set(fibers_b.keys()))
    issues = {}

    for fnum in all_fibers:
        ra = fibers_a.get(fnum)
        rb = fibers_b.get(fnum)
        a_tags, b_tags = [], []

        def _check(r, tags, pop_median_refl, dir_is_A):
            """Flag ONLY severe launch-end issues — the kind where the fiber
            silently disappears from the splice report.  We deliberately skip
            soft signals like 'NO_FIRST_SPLICE' (too noisy; many fibers have
            sub-threshold splices that don't get detected)."""
            if r is None:
                tags.append('FILE_MISSING')
                return
            launch_evt, end_km, n_events = _fiber_launch_info(r)

            # No events at all — fiber is completely silent
            if n_events == 0:
                tags.append('NO_EVENTS')
                return

            # Launch-event loss check — signed comparison (no abs()).
            # Healthy MFD-mismatch launches show a strong gainer at or below
            # -0.5 dB (more negative).  Strict greater-than: a launch_loss
            # of EXACTLY -0.5 dB is good (no flag); -0.49 dB is bad; -0.51
            # dB is good.
            if launch_evt is not None:
                launch_loss_signed = launch_evt.get('splice_loss') or 0.0
                if launch_loss_signed > hi_loss:    # default hi_loss = -0.5
                    tags.append(f'LAUNCH_LOSS{launch_loss_signed:+.2f}dB')
                refl = launch_evt.get('reflection') or 0.0
                if refl > bad_refl:
                    tags.append(f'BAD_LAUNCH_REFL{refl:+.1f}dB')

        _check(ra, a_tags, a_refl_median, dir_is_A=True)
        _check(rb, b_tags, b_refl_median, dir_is_A=False)

        if not a_tags and not b_tags:
            continue

        # Severity: HIGH for immediate-end / no-events / high-launch-loss,
        # REVIEW for missing-file / bad-refl, WATCH for only outlier / no-first.
        all_tags = a_tags + b_tags
        is_high = any(t.startswith(('NO_EVENTS',
                                    'HIGH_LAUNCH_LOSS', 'FILE_MISSING'))
                      for t in all_tags)
        is_review = any(t.startswith(('BAD_LAUNCH_REFL',)) for t in all_tags)
        severity = 'HIGH' if is_high else ('REVIEW' if is_review else 'WATCH')

        # Build a compact one-line summary (first 1–2 issue tags)
        primary = a_tags[0] if a_tags else (b_tags[0] if b_tags else '')
        dir_label = 'A' if a_tags else 'B'
        summary = f"{fnum} LAUNCH({dir_label}) {primary}"

        issues[fnum] = {
            'a_tags': a_tags,
            'b_tags': b_tags,
            'severity': severity,
            'summary': summary,
        }

    return issues


# ═══════════════════════════════════════════════════════════════════════
#  STEP 3 — Pass 1: Standard splice report analysis
#           (identical logic to splice_report_generator.py, plus A-only flagging)
# ═══════════════════════════════════════════════════════════════════════

def analyze_all(fibers_a, fibers_b, splices, threshold,
                bend_threshold=None, closure_match_km=None, **_ignored):
    """
    Pass 1: For each fiber at each known splice closure position:
      - Find A event → find matching B event → compute bidir loss → flag if above threshold
      - If no B match: flag A-only if A loss >= threshold (new vs original splice report)
      - Detect broke fibers and B-fill past breaks (same as original)

    event_source field:
      'bidir'  — both A and B direction saw it (standard splice)
      'a_only' — only A direction, no B match
      'broke'  — fiber terminates mid-span
      'bfill'  — B-direction fill past a break

    Optional kwargs (wired from the Streamlit sidebar):
      bend_threshold    — overrides BEND_THRESHOLD for this call
      closure_match_km  — overrides CLOSURE_MATCH_KM for this call
    Additional unknown kwargs are accepted and ignored for forward-compat.
    """
    bend_threshold = BEND_THRESHOLD if bend_threshold is None else float(bend_threshold)
    closure_match_km = CLOSURE_MATCH_KM if closure_match_km is None else float(closure_match_km)
    results = {}

    # End-of-fiber distances for broke detection
    eof_a = {}
    for fnum, r in fibers_a.items():
        end = [e for e in r['events'] if e['is_end']]
        eof_a[fnum] = end[0]['dist_km'] if end else 999

    # Auto-detect span: top 25% median of all EOL distances
    eof_a_vals = sorted(eof_a.values())
    if eof_a_vals:
        top_quarter_a = eof_a_vals[int(len(eof_a_vals) * 0.75):]
        total_span_a = np.median(top_quarter_a)
    else:
        total_span_a = 0

    eof_b = {}
    for fnum, r in fibers_b.items():
        end = [e for e in r['events'] if e['is_end']]
        eof_b[fnum] = end[0]['dist_km'] if end else 999

    eof_b_vals = sorted([v for v in eof_b.values() if v < 999])
    if eof_b_vals:
        top_quarter_b = eof_b_vals[int(len(eof_b_vals) * 0.75):]
        total_span_b = np.median(top_quarter_b)
    else:
        total_span_b = 0

    for fnum, r in fibers_a.items():
        rb = fibers_b.get(fnum)
        b_span = None
        if rb:
            b_end = [e for e in rb['events'] if e['is_end']]
            b_span = b_end[0]['dist_km'] if b_end else total_span_b

        # ── Per-fiber B-fill coverage / dead-zone pre-compute ──
        # If this fiber is A-broken and B also has a premature end/break,
        # there may be a stretch of cable that neither trace could see.
        _fiber_end_a = eof_a.get(fnum, 999)
        _fiber_is_a_broken = (0 < _fiber_end_a < total_span_a - END_REGION_KM)
        # B-fill reach = nearest-to-A-launch A-frame km that the B trace
        # can see.  B fiber sees from B-launch (A-frame = total_span_a) back
        # through b_span of fiber, so the furthest-back A-frame position it
        # reaches is (total_span_a - b_span).
        _b_fill_reach_km = None
        _dead_zone = None  # (lo_km, hi_km) in A-frame, both ends inclusive
        if _fiber_is_a_broken and b_span and total_span_a:
            _b_fill_reach_km = max(0.0, total_span_a - b_span)
            # If B's reach stops SHORT of the A-break (i.e., reach > break),
            # there's a gap between them that neither trace saw.
            if _b_fill_reach_km > _fiber_end_a + 0.2:
                _dead_zone = (_fiber_end_a, _b_fill_reach_km)

        for si, sp in enumerate(splices):
            sp_km = sp['position_km']
            # A column may be a real splice ('splice') or a bend / damage zone
            # ('bend' / 'damage').  In a phantom column, every qualifying A event
            # is treated as a bend (never a reburn) and never gets a BEND prefix
            # / offset annotation in its label — the column header already
            # tells the tech what the zone is.
            _column_kind = sp.get('column_kind', 'splice')
            _is_phantom_column = _column_kind in ('bend', 'damage')

            # ── Broke detection ──
            fiber_end = eof_a[fnum]
            a_plus_b = fiber_end + eof_b.get(fnum, 0) if fnum in eof_b else 0
            is_mid_span_break = (a_plus_b > 0 and
                                 abs(a_plus_b - total_span_a) < 3.0 and
                                 fiber_end < total_span_a - END_REGION_KM)

            if is_mid_span_break:
                # Mark as BROKE at the nearest splice to where it terminated
                nearest_splice = min(range(len(splices)),
                                     key=lambda i: abs(splices[i]['position_km'] - fiber_end))
                nearest_dist = abs(splices[nearest_splice]['position_km'] - fiber_end)
                if nearest_splice == si and nearest_dist < 2.0:
                    # Enrich label with B-fill coverage / dead-zone range
                    if _dead_zone is not None:
                        _broke_label = (f"{fnum} broke@{fiber_end:.1f}k | "
                                        f"DZ {_dead_zone[0]:.1f}-"
                                        f"{_dead_zone[1]:.1f}k")
                    elif _b_fill_reach_km is not None:
                        _broke_label = f"{fnum} broke@{fiber_end:.1f}k (B-fill OK)"
                    else:
                        _broke_label = f"{fnum} broke"
                    results[(fnum, si)] = {
                        'fiber': fnum, 'splice_idx': si,
                        'bidir_loss': None, 'a_loss': None, 'b_loss': None,
                        'bidir_dist': fiber_end,
                        'is_break': False, 'is_broke': True, 'is_bend': False,
                        'is_bfill': False, 'is_dead_zone': False,
                        'is_a_only': False, 'is_b_only': False,
                        'is_flagged': True, 'event_source': 'broke',
                        'event_type': 'BROKE', 'label': _broke_label,
                        'dead_zone_km': _dead_zone,
                        'b_fill_reach_km': _b_fill_reach_km,
                    }
                # B-fill for splices past the break
                elif sp_km > fiber_end and rb and b_span:
                    b_evt = None
                    for e in rb['events']:
                        if e['dist_km'] < 1.0 or e['is_end']: continue
                        ef_from_a = b_span - e['dist_km']
                        if abs(ef_from_a - sp_km) < POSITION_TOL:
                            if b_evt is None or abs(ef_from_a - sp_km) < abs((b_span - b_evt['dist_km']) - sp_km):
                                b_evt = e
                    if b_evt is not None:
                        b_loss_val = abs(b_evt['splice_loss'])
                        if b_loss_val >= threshold:
                            loss_str = f"{b_loss_val:.3f}"
                            if loss_str.startswith('0.'): loss_str = loss_str[1:]
                            results[(fnum, si)] = {
                                'fiber': fnum, 'splice_idx': si,
                                'bidir_loss': b_loss_val, 'a_loss': None,
                                'b_loss': b_evt['splice_loss'],
                                'bidir_dist': b_span - b_evt['dist_km'],
                                'is_break': False, 'is_broke': False, 'is_bend': False,
                                'is_bfill': True, 'is_dead_zone': False,
                                'is_a_only': False, 'is_b_only': False,
                                'is_flagged': True, 'event_source': 'bfill',
                                'event_type': b_evt['type'],
                                'label': f"{fnum} {loss_str} (B)",
                            }
                    elif (_dead_zone is not None and
                          _dead_zone[0] < sp_km < _dead_zone[1]):
                        # Column falls inside the dead zone for this fiber —
                        # neither A (past its break) nor B (past its own
                        # break) can see it.  Mark so the tech knows this
                        # splice was unobservable for this fiber.
                        results[(fnum, si)] = {
                            'fiber': fnum, 'splice_idx': si,
                            'bidir_loss': None, 'a_loss': None, 'b_loss': None,
                            'bidir_dist': sp_km,
                            'is_break': False, 'is_broke': False, 'is_bend': False,
                            'is_bfill': False, 'is_dead_zone': True,
                            'is_a_only': False, 'is_b_only': False,
                            'is_flagged': False, 'event_source': 'dead_zone',
                            'event_type': 'DEAD_ZONE',
                            'label': f"{fnum} DZ",
                            'dead_zone_km': _dead_zone,
                            'b_fill_reach_km': _b_fill_reach_km,
                        }
                continue

            # ── Find A event near this splice ──
            ea = None
            for e in r['events']:
                if abs(e['dist_km'] - sp_km) < POSITION_TOL and e['dist_km'] > 1.0 and not e['is_end']:
                    if ea is None or abs(e['dist_km'] - sp_km) < abs(ea['dist_km'] - sp_km):
                        ea = e

            if ea is None:
                continue

            # ── Find matching B event ──
            eb = None
            b_loss = None
            b_from_a = None
            if rb and b_span:
                for e in rb['events']:
                    if e['dist_km'] < 1.0 or e['is_end']: continue
                    ef_from_a = b_span - e['dist_km']
                    if abs(ef_from_a - ea['dist_km']) < POSITION_TOL:
                        if eb is None or abs(ef_from_a - ea['dist_km']) < abs((b_span - eb['dist_km']) - ea['dist_km']):
                            eb = e
                            b_loss = e['splice_loss']
                            b_from_a = ef_from_a

            # ── A event but no B event in table ──
            # Try to measure the B-direction loss directly from the B trace
            # using wide-LSA (EXFO's "grey value" approach).  Convert the
            # splice position to the B-frame (B_dist = B_span - sp_km).
            if b_loss is None:
                a_loss_abs = abs(ea['splice_loss'])
                b_grey = None
                if rb is not None and b_span:
                    b_frame_km = b_span - sp_km
                    b_grey = _grey_loss(rb, b_frame_km)

                if b_grey is not None:
                    # Real bidirectional average using measured B grey
                    true_bidir = round((ea['splice_loss'] + b_grey) / 2.0, 4)
                    closure_center_km = sp.get('position_km_refined', sp_km)
                    is_bend_offset = _is_bend_event(ea['dist_km'], closure_center_km, true_bidir)
                    is_bend = is_bend_offset or _is_phantom_column

                    if abs(true_bidir) >= threshold or is_bend:
                        loss_str = _format_loss(true_bidir)
                        if is_bend and not _is_phantom_column:
                            offset_m = round((ea['dist_km'] - closure_center_km) * 1000, 0)
                            label = f"{fnum} BEND {loss_str} ({offset_m:+.0f}m)"
                        else:
                            label = f"{fnum} {loss_str}"
                        results[(fnum, si)] = {
                            'fiber': fnum, 'splice_idx': si,
                            'bidir_loss': true_bidir,
                            'a_loss': ea['splice_loss'], 'b_loss': b_grey,
                            'bidir_dist': ea['dist_km'],
                            'is_break': False, 'is_broke': False, 'is_bend': is_bend,
                            'is_bfill': False, 'is_a_only': False, 'is_b_only': False,
                            'is_flagged': True,
                            'event_source': 'bend' if is_bend else 'bidir_grey_b',
                            'bend_severity': _bend_severity(true_bidir) if is_bend else None,
                            'closure_offset_m': round((ea['dist_km'] - closure_center_km) * 1000, 1) if is_bend else None,
                            'event_type': ea['type'],
                            'label': label,
                            '_b_is_grey': not is_bend,
                        }
                        continue

                    # Below threshold — skip
                    continue

                # No JSON trace available — fall back to conservative (A alone) check:
                # flag as A-only if the single-direction loss alone clears threshold
                if a_loss_abs >= threshold:
                    loss_str = _format_loss(a_loss_abs)
                    closure_center_km = sp.get('position_km_refined', sp_km)
                    is_bend_offset = _is_bend_event(ea['dist_km'], closure_center_km, ea['splice_loss'])
                    is_bend = is_bend_offset or _is_phantom_column
                    if is_bend and not _is_phantom_column:
                        offset_m = round((ea['dist_km'] - closure_center_km) * 1000, 0)
                        label = f"{fnum} BEND {loss_str}(A) ({offset_m:+.0f}m)"
                    elif is_bend:
                        label = f"{fnum} {loss_str}(A)"
                    else:
                        label = f"{fnum} {loss_str} (A)"
                    results[(fnum, si)] = {
                        'fiber': fnum, 'splice_idx': si,
                        'bidir_loss': None, 'a_loss': ea['splice_loss'], 'b_loss': None,
                        'bidir_dist': ea['dist_km'],
                        'is_break': False, 'is_broke': False, 'is_bend': is_bend,
                        'is_bfill': False,
                        'is_a_only': not is_bend, 'is_b_only': False,
                        'is_flagged': True,
                        'event_source': 'bend' if is_bend else 'a_only',
                        'bend_severity': _bend_severity(ea['splice_loss']) if is_bend else None,
                        'closure_offset_m': round((ea['dist_km'] - closure_center_km) * 1000, 1) if is_bend else None,
                        'event_type': ea['type'],
                        'label': label,
                    }
                continue

            # ── A+B bidirectional ──
            bidir_loss = round((ea['splice_loss'] + b_loss) / 2.0, 4)
            bidir_dist = round((ea['dist_km'] + b_from_a) / 2.0, 4)

            is_reflective = ea['type'].startswith('1F')
            has_weak_fresnel = ea['reflection'] < -30.0
            is_break = is_reflective and has_weak_fresnel and ea['dist_km'] < (total_span_a - END_REGION_KM)

            # ── BEND check (ZeroDBIFTHEN Flag-3 rule) ──
            # If the event position is offset from the true closure center
            # by more than CLOSURE_MATCH_KM (150 m), this is a BEND not a
            # splice reburn.  Use the refined (mode-based) center, falling
            # back to the coarse position_km if refinement hasn't run.
            closure_center_km = sp.get('position_km_refined', sp_km)
            is_bend_offset = _is_bend_event(ea['dist_km'], closure_center_km, bidir_loss)
            # Phantom columns always classify as bends (unless they're breaks).
            # The cell label stays clean — column header carries the zone type.
            is_bend = (not is_break) and (is_bend_offset or _is_phantom_column)

            is_flagged = (abs(bidir_loss) >= threshold) or is_break or is_bend
            if not is_flagged:
                continue

            if is_break:
                offset_m = round((bidir_dist - sp_km) * 1000, 1)
                uni_loss = abs(ea['splice_loss'])
                refl_db = ea['reflection']
                refl_str = f" {uni_loss:.3f} uni reflection {refl_db:.0f}"
                if refl_db > -35.0:
                    break_type = " air gap"
                else:
                    break_type = ""
                label = f"{fnum} BREAK {bidir_loss:.3f} ({abs(offset_m):.0f}m from splice){refl_str}{break_type}"
            elif is_bend:
                loss_str = _format_loss(bidir_loss)
                if _is_phantom_column:
                    # Column already says 'bends@X km' — cell just shows fiber + loss
                    label = f"{fnum} {loss_str}"
                else:
                    offset_m = round((ea['dist_km'] - closure_center_km) * 1000, 0)
                    label = f"{fnum} BEND {loss_str} ({offset_m:+.0f}m)"
            else:
                loss_str = _format_loss(bidir_loss)
                label = f"{fnum} {loss_str}"

            results[(fnum, si)] = {
                'fiber': fnum, 'splice_idx': si,
                'bidir_loss': bidir_loss,
                'a_loss': ea['splice_loss'], 'b_loss': b_loss,
                'bidir_dist': bidir_dist,
                'is_break': is_break, 'is_broke': False, 'is_bend': is_bend,
                'is_bfill': False, 'is_a_only': False, 'is_b_only': False,
                'is_flagged': True,
                'event_source': 'bend' if is_bend else 'bidir',
                'bend_severity': _bend_severity(bidir_loss) if is_bend else None,
                'closure_offset_m': round((ea['dist_km'] - closure_center_km) * 1000, 1) if is_bend else None,
                'event_type': ea['type'],
                'label': label,
                'fresnel': ea['reflection'] if is_reflective else None,
            }

    return results


# ═══════════════════════════════════════════════════════════════════════
#  STEP 4 — Pass 2: Scan all B-direction events not caught in Pass 1
# ═══════════════════════════════════════════════════════════════════════

def scan_b_events(fibers_a, fibers_b, splices, threshold, existing_results, total_span_a,
                  bend_threshold=None, closure_match_km=None, **_ignored):
    """
    Pass 2: For every B-direction event above threshold that was NOT already
    caught in Pass 1, find the nearest splice position (within 1.5 km) and report it.

    This is how EXFO finds events like fiber 325's 0.340 dB entry that only
    exists in the B-direction event table with no matching A-direction event.

    Returns a dict of (fnum, si) -> result — same structure as analyze_all().
    Does NOT overwrite any existing_results entries.
    """
    new_results = {}

    for fnum, rb in fibers_b.items():
        ra = fibers_a.get(fnum)

        # B-direction span (EOL)
        b_end_events = [e for e in rb['events'] if e['is_end']]
        if not b_end_events:
            continue
        b_span = b_end_events[0]['dist_km']

        # A-direction EOL (to know if this fiber is broken)
        ra_end_km = total_span_a
        if ra:
            a_end = [e for e in ra['events'] if e['is_end']]
            if a_end:
                ra_end_km = a_end[0]['dist_km']

        for e in rb['events']:
            if e['dist_km'] < 1.0 or e['is_end']:
                continue

            b_loss_signed = e['splice_loss']
            b_loss_abs = abs(b_loss_signed)
            # Gate: skip clearly-too-small B events.  Use B alone (not B/2)
            # because the real bidir depends on the A grey value we haven't
            # measured yet.  Anything with single-dir loss below threshold
            # can't possibly produce a bidir above threshold unless A grey
            # is even larger, which is unlikely.
            if b_loss_abs < threshold * 0.75:
                continue

            # Convert B-frame position to A-frame
            a_frame_km = b_span - e['dist_km']
            if a_frame_km < 0.5:
                continue  # launch artifact near B-end

            # Find nearest splice position within tolerance
            nearest_si = None
            nearest_dist = float('inf')
            for si, sp in enumerate(splices):
                d = abs(sp['position_km'] - a_frame_km)
                if d < nearest_dist:
                    nearest_dist = d
                    nearest_si = si

            if nearest_si is None or nearest_dist > POSITION_TOL:
                continue  # not near any known splice position

            # Already caught by Pass 1?
            if (fnum, nearest_si) in existing_results:
                continue

            # Already found a better match in this pass?
            if (fnum, nearest_si) in new_results:
                existing_a_frame = new_results[(fnum, nearest_si)]['bidir_dist']
                if nearest_dist >= abs(splices[nearest_si]['position_km'] - existing_a_frame):
                    continue

            # Look for A-direction event near the same A-frame position
            a_evt = None
            if ra:
                for ae in ra['events']:
                    if ae['dist_km'] < 1.0 or ae['is_end']: continue
                    if abs(ae['dist_km'] - a_frame_km) < POSITION_TOL:
                        if a_evt is None or abs(ae['dist_km'] - a_frame_km) < abs(a_evt['dist_km'] - a_frame_km):
                            a_evt = ae

            closure_center_km = splices[nearest_si].get('position_km_refined',
                                                         splices[nearest_si]['position_km'])

            if a_evt is not None:
                # A event exists — compute bidirectional
                bidir = round((a_evt['splice_loss'] + b_loss_signed) / 2.0, 4)
                is_bend = _is_bend_event(a_frame_km, closure_center_km, bidir)
                if abs(bidir) < threshold and not is_bend:
                    continue
                loss_str = _format_loss(bidir)
                if is_bend:
                    offset_m = round((a_frame_km - closure_center_km) * 1000, 0)
                    label = f"{fnum} BEND {loss_str} ({offset_m:+.0f}m)"
                else:
                    label = f"{fnum} {loss_str}"
                new_results[(fnum, nearest_si)] = {
                    'fiber': fnum, 'splice_idx': nearest_si,
                    'bidir_loss': bidir,
                    'a_loss': a_evt['splice_loss'], 'b_loss': b_loss_signed,
                    'bidir_dist': a_frame_km,
                    'is_break': False, 'is_broke': False, 'is_bend': is_bend,
                    'is_bfill': False, 'is_a_only': False, 'is_b_only': False,
                    'is_flagged': True,
                    'event_source': 'bend' if is_bend else 'bidir',
                    'bend_severity': _bend_severity(bidir) if is_bend else None,
                    'closure_offset_m': round((a_frame_km - closure_center_km) * 1000, 1) if is_bend else None,
                    'event_type': a_evt['type'],
                    'label': label,
                }
            else:
                # B event but no A event in table — measure the A-direction
                # loss at this position from the A JSON trace (grey value).
                a_grey = _grey_loss(ra, a_frame_km) if ra is not None else None

                if a_grey is not None:
                    true_bidir = round((a_grey + b_loss_signed) / 2.0, 4)
                    is_bend = _is_bend_event(a_frame_km, closure_center_km, true_bidir)
                    if abs(true_bidir) < threshold and not is_bend:
                        continue
                    loss_str = _format_loss(true_bidir)
                    if is_bend:
                        offset_m = round((a_frame_km - closure_center_km) * 1000, 0)
                        label = f"{fnum} BEND {loss_str} ({offset_m:+.0f}m)"
                    else:
                        label = f"{fnum} {loss_str}"
                    new_results[(fnum, nearest_si)] = {
                        'fiber': fnum, 'splice_idx': nearest_si,
                        'bidir_loss': true_bidir,
                        'a_loss': a_grey, 'b_loss': b_loss_signed,
                        'bidir_dist': a_frame_km,
                        'is_break': False, 'is_broke': False, 'is_bend': is_bend,
                        'is_bfill': False, 'is_a_only': False, 'is_b_only': False,
                        'is_flagged': True,
                        'event_source': 'bend' if is_bend else 'bidir_grey_a',
                        'bend_severity': _bend_severity(true_bidir) if is_bend else None,
                        'closure_offset_m': round((a_frame_km - closure_center_km) * 1000, 1) if is_bend else None,
                        'event_type': e['type'],
                        'label': label,
                        '_a_is_grey': not is_bend,
                    }
                    continue

                # No JSON trace — fall back to single-direction check (B alone
                # ≥ threshold, which is conservative but honest)
                if b_loss_abs >= threshold:
                    is_bend = _is_bend_event(a_frame_km, closure_center_km, b_loss_signed)
                    loss_str = _format_loss(b_loss_abs)
                    if is_bend:
                        offset_m = round((a_frame_km - closure_center_km) * 1000, 0)
                        label = f"{fnum} BEND {loss_str}(B) ({offset_m:+.0f}m)"
                    else:
                        label = f"{fnum} {loss_str} (B)"
                    new_results[(fnum, nearest_si)] = {
                        'fiber': fnum, 'splice_idx': nearest_si,
                        'bidir_loss': None,
                        'a_loss': None, 'b_loss': b_loss_signed,
                        'bidir_dist': a_frame_km,
                        'is_break': False, 'is_broke': False, 'is_bend': is_bend,
                        'is_bfill': False,
                        'is_a_only': False, 'is_b_only': not is_bend,
                        'is_flagged': True,
                        'event_source': 'bend' if is_bend else 'b_only',
                        'bend_severity': _bend_severity(b_loss_signed) if is_bend else None,
                        'closure_offset_m': round((a_frame_km - closure_center_km) * 1000, 1) if is_bend else None,
                        'event_type': e['type'],
                        'label': label,
                    }

    return new_results


# ═══════════════════════════════════════════════════════════════════════
#  STEP 4b — APRIL 23 revision: A-first standalone-event classifier
#  (bends / breaks that are not at any VALID splice closure)
# ═══════════════════════════════════════════════════════════════════════

def scan_a_standalone_events(fibers_a, splices, existing_results, total_span_a,
                             bend_threshold=None, closure_match_km=None,
                             nearest_tol_km=2.0):
    """Every A-direction non-end event that was NOT covered by Pass 1 gets
    classified as a BEND or a BREAK/BROKE.  This replaces the old behaviour
    where events inside a phantom splice column would render as splice
    reburns even though they were actually bends.

    Rule (from the tech's April 23 instructions):
      • We build the template off A direction events.
      • Any A events NOT at a splice column → bend or break.

    Returns dict (fnum, synthetic_si) → result-dict.  synthetic_si is the
    index of the NEAREST valid closure (used only for ribbon-grid placement;
    the event is displayed with a BEND/BREAK label + offset annotation)."""
    bt = BEND_THRESHOLD   if bend_threshold   is None else bend_threshold
    cm = CLOSURE_MATCH_KM if closure_match_km is None else closure_match_km

    new_results = {}
    seen_keys = set(existing_results.keys())

    # Pre-compute the refined center of every closure for fast nearest lookup
    closure_centers = [(si, sp.get('position_km_refined', sp['position_km']))
                       for si, sp in enumerate(splices)]

    for fnum, ra in fibers_a.items():
        # Skip fibers that are broke — broke fibers get special treatment
        events = ra.get('events') or []
        end_events = [e for e in events if e.get('is_end')]
        if not end_events:
            continue
        eof_a = end_events[0]['dist_km']
        is_broken = eof_a < (total_span_a - END_REGION_KM)

        for e in events:
            if e['is_end']:
                continue
            if e['dist_km'] < 1.0:
                continue  # launch region — handled separately
            loss = e.get('splice_loss') or 0.0
            if abs(loss) < bt:
                continue

            # Find nearest closure (by refined center)
            best_si, best_d = None, float('inf')
            for si, c in closure_centers:
                d = abs(e['dist_km'] - c)
                if d < best_d:
                    best_d = d
                    best_si = si

            # If this event IS within CLOSURE_MATCH_KM of a real closure,
            # analyze_all (Pass 1) already handled it — skip.
            if best_d <= cm:
                continue
            # If very far from every closure, still bucket to the nearest
            # one for display purposes (within nearest_tol_km).  Farther than
            # that means the event lives between closures too far from any;
            # still assign to the nearest column for visibility.

            # Already flagged by Pass 1 / Pass 2 at this closure position?
            key = (fnum, best_si)
            if key in seen_keys or key in new_results:
                continue

            is_reflective = e.get('is_reflective') or str(e.get('type','')).startswith('1F')
            refl = e.get('reflection') or 0.0
            has_weak_fresnel = refl < -30.0

            # BREAK: reflective, has weak Fresnel, mid-span
            if is_reflective and has_weak_fresnel and e['dist_km'] < (total_span_a - END_REGION_KM):
                loss_str = _format_loss(loss)
                refl_str = f" {abs(loss):.3f} uni reflection {refl:.0f}"
                break_type = " air gap" if refl > -35.0 else ""
                label = (f"{fnum} BREAK {loss_str} @ {e['dist_km']:.3f}km"
                         f"{refl_str}{break_type}")
                new_results[key] = {
                    'fiber': fnum, 'splice_idx': best_si,
                    'bidir_loss': loss, 'a_loss': loss, 'b_loss': None,
                    'bidir_dist': e['dist_km'],
                    'is_break': True, 'is_broke': False, 'is_bend': False,
                    'is_bfill': False, 'is_a_only': False, 'is_b_only': False,
                    'is_flagged': True, 'event_source': 'break_standalone',
                    'event_type': e['type'],
                    'label': label,
                    'fresnel': refl,
                }
                continue

            # BEND: everything else above threshold.  If the nearest column
            # is itself a phantom bend/damage zone, the column header already
            # describes the zone — keep the cell label clean.
            target_sp = splices[best_si]
            target_is_phantom = target_sp.get('column_kind') in ('bend', 'damage')
            loss_str = _format_loss(loss)
            if target_is_phantom:
                label = f"{fnum} {loss_str}"
            else:
                offset_m = round((e['dist_km'] - closure_centers[best_si][1]) * 1000, 0)
                label = f"{fnum} BEND {loss_str} ({offset_m:+.0f}m)"
            new_results[key] = {
                'fiber': fnum, 'splice_idx': best_si,
                'bidir_loss': loss, 'a_loss': loss, 'b_loss': None,
                'bidir_dist': e['dist_km'],
                'is_break': False, 'is_broke': False, 'is_bend': True,
                'is_bfill': False, 'is_a_only': False, 'is_b_only': False,
                'is_flagged': True, 'event_source': 'bend_standalone',
                'bend_severity': _bend_severity(loss),
                'closure_offset_m': float(offset_m),
                'event_type': e['type'],
                'label': label,
            }
    return new_results


# ═══════════════════════════════════════════════════════════════════════
#  STEP 4c — APRIL 23 revision: restricted past-break B-fill scan
#  (replaces the blanket Pass-2 B scan with a targeted past-break pass)
# ═══════════════════════════════════════════════════════════════════════

def scan_b_past_breaks(fibers_a, fibers_b, splices, threshold, existing_results,
                       total_span_a):
    """For fibers that are BROKE on the A side, scan the B direction in the
    zone PAST the A-side break (so closer to the B-end than the break
    position) for events we haven't seen yet — these populate B-fill cells.

    This is the narrower B-direction usage the tech asked for: 'We use B
    direction only to look after the breaks on A sides to see if there are
    other events we are missing.'"""
    new_results = {}
    seen_keys = set(existing_results.keys())

    # Cache A break positions (km, in A-frame)
    a_break_km = {}
    for fnum, ra in fibers_a.items():
        end = [e for e in ra['events'] if e.get('is_end')]
        if not end:
            continue
        eof = end[0]['dist_km']
        if eof < total_span_a - END_REGION_KM:
            a_break_km[fnum] = eof

    if not a_break_km:
        return new_results

    for fnum, brk_km in a_break_km.items():
        rb = fibers_b.get(fnum)
        if rb is None:
            continue
        b_end_events = [e for e in rb['events'] if e.get('is_end')]
        if not b_end_events:
            continue
        b_span = b_end_events[0]['dist_km']

        # Scan B events whose A-frame position is GREATER than brk_km
        for e in rb['events']:
            if e.get('is_end'):
                continue
            if e['dist_km'] < 1.0:
                continue
            a_frame = b_span - e['dist_km']
            if a_frame <= brk_km + 0.2:   # 200m buffer past the break
                continue
            b_loss = e.get('splice_loss') or 0.0
            if abs(b_loss) < threshold:
                continue

            # Find nearest splice position (A-frame)
            nearest_si, nearest_d = None, float('inf')
            for si, sp in enumerate(splices):
                c = sp.get('position_km_refined', sp['position_km'])
                d = abs(a_frame - c)
                if d < nearest_d:
                    nearest_d = d
                    nearest_si = si
            if nearest_si is None:
                continue
            # Keep loose tolerance for B-fill (the whole point is bridging a gap)
            if nearest_d > POSITION_TOL:
                continue

            key = (fnum, nearest_si)
            if key in seen_keys or key in new_results:
                continue

            loss_str = _format_loss(b_loss)
            new_results[key] = {
                'fiber': fnum, 'splice_idx': nearest_si,
                'bidir_loss': abs(b_loss), 'a_loss': None,
                'b_loss': b_loss, 'bidir_dist': a_frame,
                'is_break': False, 'is_broke': False, 'is_bend': False,
                'is_bfill': True,
                'is_a_only': False, 'is_b_only': False,
                'is_flagged': True, 'event_source': 'bfill',
                'event_type': e['type'],
                'label': f"{fnum} {loss_str} (B-fill)",
            }

    return new_results


# ═══════════════════════════════════════════════════════════════════════
#  STEP 5 — Group into ribbons and build cell values
# ═══════════════════════════════════════════════════════════════════════

def build_ribbon_data(results, n_fibers, ribbon_size, n_splices, launch_issues=None):
    """Group flagged events into ribbon rows × splice columns.  If
    launch_issues is provided, each ribbon gets an extra 'launch_cell' entry
    summarising which of its fibers have launch-end issues — the write_xlsx
    function renders these into the ILA:A column."""
    n_ribbons = (n_fibers + ribbon_size - 1) // ribbon_size
    grid = {}

    for (fnum, si), res in results.items():
        ri = (fnum - 1) // ribbon_size
        key = (ri, si)
        if key not in grid:
            grid[key] = []
        grid[key].append(res)

    cells = {}
    for (ri, si), res_list in grid.items():
        res_list.sort(key=lambda x: x['fiber'])

        # Group fibers with same loss and same source type
        groups = []
        for res in res_list:
            merged = False
            for g in groups:
                if (res['bidir_loss'] is not None and g['loss'] is not None and
                        abs(res['bidir_loss'] - g['loss']) < 0.002 and
                        not res['is_break'] and not res['is_broke'] and
                        not res.get('is_bend', False) and not g.get('is_bend', False) and
                        not g['is_break'] and not g['is_broke'] and
                        res.get('event_source') == g.get('event_source')):
                    g['fibers'].append(res['fiber'])
                    merged = True
                    break
            if not merged:
                groups.append({
                    'fibers': [res['fiber']],
                    'loss': res['bidir_loss'],
                    'is_break': res['is_break'],
                    'is_broke': res['is_broke'],
                    'is_bend':  res.get('is_bend', False),
                    'is_bfill': res.get('is_bfill', False),
                    'is_dead_zone': res.get('is_dead_zone', False),
                    'is_gainer': res.get('is_gainer', False),
                    'is_a_only': res.get('is_a_only', False),
                    'is_b_only': res.get('is_b_only', False),
                    'event_source': res.get('event_source', 'bidir'),
                    'label': res['label'],
                    'res': res,
                })

        # Build cell text — label shows source for A-only and B-only
        parts = []
        for g in groups:
            if g.get('is_dead_zone'):
                # Collapse multi-fiber dead zones into "F1,F2,... DZ"
                fib_str = ','.join(str(f) for f in g['fibers'])
                parts.append(f"{fib_str} DZ")
            elif g['is_broke']:
                # Use the enriched broke label (includes position + DZ range)
                parts.append(g['label'])
            elif g['is_break']:
                parts.append(g['label'])
            elif g.get('is_bend'):
                # Use the full label (includes "BEND" marker and offset)
                parts.append(g['label'])
            elif g['is_a_only']:
                fib_str = ','.join(str(f) for f in g['fibers'])
                raw_loss = g['res']['a_loss']
                loss_abs = abs(raw_loss) if raw_loss is not None else 0
                loss_str = f"{loss_abs:.3f}"
                if loss_str.startswith('0.'): loss_str = loss_str[1:]
                # Show estimated bidir value (like Steven's "bidi .173" annotation)
                est_bd = g['res'].get('est_bidir')
                if est_bd is not None:
                    bd_str = f"{est_bd:.3f}"
                    if bd_str.startswith('0.'): bd_str = bd_str[1:]
                    parts.append(f"{fib_str} {loss_str} (A) bidi {bd_str}")
                else:
                    parts.append(f"{fib_str} {loss_str} (A)")
            elif g['is_b_only']:
                fib_str = ','.join(str(f) for f in g['fibers'])
                raw_loss = g['res']['b_loss']
                loss_abs = abs(raw_loss) if raw_loss is not None else 0
                loss_str = f"{loss_abs:.3f}"
                if loss_str.startswith('0.'): loss_str = loss_str[1:]
                # Show estimated bidir value
                est_bd = g['res'].get('est_bidir')
                if est_bd is not None:
                    bd_str = f"{est_bd:.3f}"
                    if bd_str.startswith('0.'): bd_str = bd_str[1:]
                    parts.append(f"{fib_str} {loss_str} (B) bidi {bd_str}")
                else:
                    parts.append(f"{fib_str} {loss_str} (B)")
            elif g.get('is_bfill'):
                fib_str = ','.join(str(f) for f in g['fibers'])
                loss = g['loss']
                loss_str = f"{loss:.3f}" if loss is not None else "?"
                if loss_str.startswith('0.'): loss_str = loss_str[1:]
                parts.append(f"{fib_str} {loss_str} (B-fill)")
            else:
                fib_str = ','.join(str(f) for f in g['fibers'])
                loss = g['loss']
                loss_str = f"{loss:.3f}" if loss is not None else "?"
                if loss_str.startswith('0.'): loss_str = loss_str[1:]
                parts.append(f"{fib_str} {loss_str}")

        cell_text = ' '.join(parts)
        is_break = any(g['is_break'] for g in groups)
        is_broke = any(g['is_broke'] for g in groups)
        is_bend  = any(g.get('is_bend', False) for g in groups)
        is_bfill = any(g.get('is_bfill', False) for g in groups)

        # Has a standard bidir reburn in this cell?
        has_standard_reburn = any(
            not g['is_break'] and not g['is_broke'] and
            not g.get('is_bend') and
            not g.get('is_bfill') and not g.get('is_a_only') and
            not g.get('is_b_only')
            for g in groups
        )
        # A-only / B-only only drive color if no higher-priority event present
        is_a_only = (any(g.get('is_a_only', False) for g in groups) and
                     not is_break and not is_broke and not is_bfill and not has_standard_reburn)
        is_b_only = (any(g.get('is_b_only', False) for g in groups) and
                     not is_break and not is_broke and not is_bfill and not has_standard_reburn)

        # If estimated bidir still clears threshold, use a stronger shade
        est_bidir_flagged = any(g['res'].get('est_bidir_flagged', False) for g in groups
                                if g.get('is_a_only') or g.get('is_b_only'))

        max_loss = max((g['loss'] for g in groups if g['loss'] is not None), default=0)

        is_dead_zone = any(g.get('is_dead_zone', False) for g in groups)
        is_gainer    = any(g.get('is_gainer', False) for g in groups)

        cells[(ri, si)] = {
            'text': cell_text,
            'is_break': is_break,
            'is_broke': is_broke,
            'is_bend':  is_bend,
            'is_bfill': is_bfill,
            'is_dead_zone': is_dead_zone,
            'is_gainer': is_gainer,
            'is_a_only': is_a_only,
            'is_b_only': is_b_only,
            'est_bidir_flagged': est_bidir_flagged,
            'max_loss': max_loss,
        }

    # ── Per-ribbon launch-issue summaries (for the ILA:A / ILA:B columns) ──
    launch_cells_a = {}   # ribbon_index → dict {text, severity}
    launch_cells_b = {}
    if launch_issues:
        per_ribbon_a = {}   # ri → list of (fnum, severity, tag)
        per_ribbon_b = {}
        for fnum, info in launch_issues.items():
            ri = (fnum - 1) // ribbon_size
            for tag in info.get('a_tags', []):
                per_ribbon_a.setdefault(ri, []).append((fnum, info['severity'], tag))
            for tag in info.get('b_tags', []):
                per_ribbon_b.setdefault(ri, []).append((fnum, info['severity'], tag))

        def _sev_order(s):
            return {'HIGH': 0, 'REVIEW': 1, 'WATCH': 2}.get(s, 3)

        for ri, items in per_ribbon_a.items():
            worst = min(items, key=lambda x: _sev_order(x[1]))[1]
            # Compact label: fiber# + abbreviated tag
            parts = [f"{f} {tag.split('@')[0].split('+')[0]}" for f, _, tag in items]
            launch_cells_a[ri] = {'text': ' '.join(parts[:6]) +
                                          (f" +{len(parts)-6} more" if len(parts) > 6 else ''),
                                   'severity': worst}
        for ri, items in per_ribbon_b.items():
            worst = min(items, key=lambda x: _sev_order(x[1]))[1]
            parts = [f"{f} {tag.split('@')[0].split('+')[0]}" for f, _, tag in items]
            launch_cells_b[ri] = {'text': ' '.join(parts[:6]) +
                                          (f" +{len(parts)-6} more" if len(parts) > 6 else ''),
                                   'severity': worst}

    return cells, launch_cells_a, launch_cells_b


# ═══════════════════════════════════════════════════════════════════════
#  STEP 6 — Generate Excel
# ═══════════════════════════════════════════════════════════════════════

def ribbon_label(ri, ribbon_size, n_fibers):
    first = ri * ribbon_size + 1
    last = min(first + ribbon_size - 1, n_fibers)
    ribbon_num = ri + 1
    tube = ''
    if ri < 48:
        tube_letter = chr(ord('A') + ri // 2)
        tube_num = (ri % 2) + 1
        tube = f" ({tube_letter}{tube_num})"
    return f"Fiber {first}-{last} ({ribbon_num}){tube}"


def write_xlsx(cells, splices, n_fibers, ribbon_size, output_path, site_a, site_b, span_km,
               launch_cells_a=None, launch_cells_b=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Splice Report"

    n_ribbons = (n_fibers + ribbon_size - 1) // ribbon_size
    n_splices = len(splices)

    # ── Styles ──
    hdr_font    = Font(bold=True, size=10, color="FFFFFF")
    hdr_fill    = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    data_font   = Font(size=8)
    ribbon_font = Font(size=9)
    a_km_font   = Font(bold=True, size=9, color="1F4E79")
    b_km_font   = Font(bold=True, size=9, color="8B0000")

    # Cell fill/font for each event type
    red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")   # A+B reburn
    break_fill  = PatternFill(start_color="FF4444", end_color="FF4444", fill_type="solid")   # break
    break_font  = Font(bold=True, size=8, color="FFFFFF")
    # April 23 revision: BROKE is now rendered the same as BREAK — both are
    # physical damage where the fiber has lost transmission.  Same red fill,
    # same white bold text, same category in the tech's eye.
    broke_fill  = PatternFill(start_color="FF4444", end_color="FF4444", fill_type="solid")   # red
    broke_font  = Font(bold=True, size=8, color="FFFFFF")
    bfill_fill  = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")   # B-fill past break
    bfill_font  = Font(size=8, color="1F4E79")
    dz_fill     = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")   # dead zone (gray)
    dz_font     = Font(size=8, italic=True, color="3F3F3F")
    gainer_fill = PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid")   # field gainer (mint green)
    gainer_font = Font(bold=True, size=8, color="1B5E20")
    aonly_fill  = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")   # A-only (light yellow, est bidir OK)
    aonly_font  = Font(size=8, color="7F6000")
    aonly_fill2 = PatternFill(start_color="FF7043", end_color="FF7043", fill_type="solid")   # A-only (coral, est bidir >= threshold) — deliberately non-yellow
    aonly_font2 = Font(bold=True, size=8, color="FFFFFF")
    bonly_fill  = PatternFill(start_color="E8D5F5", end_color="E8D5F5", fill_type="solid")   # B-only (lavender, est bidir OK)
    bonly_font  = Font(size=8, color="4B0082")
    bonly_fill2 = PatternFill(start_color="C084FC", end_color="C084FC", fill_type="solid")   # B-only (purple, est bidir >= threshold)
    bonly_font2 = Font(bold=True, size=8, color="1A0033")
    # BEND: teal / cyan — clearly distinct from splice colors so bends stand out
    # BEND cells: single yellow fill for every bend (no severity shading).
    # Matches the tech's yellow-highlight style on Cle Elum.
    bend_fill        = PatternFill(start_color="FFEB3B", end_color="FFEB3B", fill_type="solid")
    bend_font        = Font(bold=True, size=8, color="5D4037")
    # Keep the old three-name aliases pointing at the single fill so any
    # downstream reference still resolves.  bend_font_high is just bend_font.
    bend_fill_watch  = bend_fill
    bend_fill_review = bend_fill
    bend_fill_high   = bend_fill
    bend_font_high   = bend_font
    # LAUNCH ISSUE: single orange fill — warns the tech a fiber had
    # launch-end trouble (broken at launch, damaged connector, truncated
    # event table, bad reflectance).  Orange was chosen to be clearly
    # distinct from the pink A+B reburn fill.  Severity tiers were
    # collapsed per tech feedback — a launch issue is a launch issue.
    launch_fill        = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    launch_font        = Font(bold=True, size=8, color="5D2E00")
    # Back-compat aliases (in case anything else in the codebase still
    # references the per-severity names)
    launch_fill_high = launch_fill_review = launch_fill_watch = launch_fill
    launch_font_high = launch_font_watch = launch_font

    border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )

    # ── Row 1: A→B distances (km / ft) ──
    ws.cell(row=1, column=2, value="A→B:").font = a_km_font
    ws.cell(row=2, column=2, value="B→A:").font = b_km_font
    for si, sp in enumerate(splices):
        col = si + 3
        km = sp['position_km']
        ft = km * 3280.84
        b_km = span_km - km
        b_ft = b_km * 3280.84
        c1 = ws.cell(row=1, column=col, value=f"{km:.2f}km / {ft:,.0f}ft")
        c1.font = a_km_font
        c1.alignment = Alignment(horizontal='center')
        c2 = ws.cell(row=2, column=col, value=f"{b_km:.2f}km / {b_ft:,.0f}ft")
        c2.font = b_km_font
        c2.alignment = Alignment(horizontal='center')
    end_col = n_splices + 3
    ws.cell(row=1, column=end_col, value=f"{span_km:.2f}km / {span_km*3280.84:,.0f}ft").font = a_km_font
    ws.cell(row=2, column=end_col, value="0.00km / 0ft").font = b_km_font

    # ── Row 3: Headers ──
    ws.cell(row=3, column=1, value="Ribbon").font = hdr_font
    ws.cell(row=3, column=1).fill = hdr_fill
    ws.cell(row=3, column=2, value=f"ILA:{site_a}").font = hdr_font
    ws.cell(row=3, column=2).fill = hdr_fill
    # Alternate fill colors for phantom (bend / damage) column headers so they
    # stand out from the blue splice headers at a glance.
    # Bend-column header: same yellow as the bend cells beneath it.
    # Damage-column header: red (matches the red break/broke cells beneath).
    hdr_fill_bend   = PatternFill(start_color="FFEB3B", end_color="FFEB3B", fill_type="solid")
    hdr_fill_damage = PatternFill(start_color="FF4444", end_color="FF4444", fill_type="solid")
    for si, sp in enumerate(splices):
        col = si + 3
        kind = sp.get('column_kind', 'splice')
        if kind == 'bend':
            ref_km = sp.get('position_km_refined', sp['position_km'])
            header = f"Bends @ {ref_km:.2f}km"
            cell = ws.cell(row=3, column=col, value=header)
            cell.fill = hdr_fill_bend
        elif kind == 'damage':
            ref_km = sp.get('position_km_refined', sp['position_km'])
            header = f"Damage @ {ref_km:.2f}km"
            cell = ws.cell(row=3, column=col, value=header)
            cell.fill = hdr_fill_damage
        else:
            disp_n = sp.get('splice_display_num', si + 1)
            cell = ws.cell(row=3, column=col, value=f"Splice {disp_n}")
            cell.fill = hdr_fill
        cell.font = hdr_font
    ws.cell(row=3, column=end_col, value=f"ILA:{site_b}").font = hdr_font
    ws.cell(row=3, column=end_col).fill = hdr_fill

    # ── Data rows ──
    def _launch_fill(sev):
        # Single light-pink tier for all launch issues (severity ignored)
        return launch_fill, launch_font

    for ri in range(n_ribbons):
        row = ri + 4
        ws.cell(row=row, column=1, value=ribbon_label(ri, ribbon_size, n_fibers)).font = ribbon_font

        # ── ILA:A column (col 2) — launch-issue summary for A direction ──
        ila_a_cell = ws.cell(row=row, column=2)
        ila_a_cell.border = border
        ila_a_cell.alignment = Alignment(wrap_text=True, vertical='center')
        if launch_cells_a and ri in launch_cells_a:
            lc = launch_cells_a[ri]
            ila_a_cell.value = lc['text']
            f, fn = _launch_fill(lc['severity'])
            ila_a_cell.fill = f
            ila_a_cell.font = fn

        # ── ILA:B column (end_col) — launch-issue summary for B direction ──
        ila_b_cell = ws.cell(row=row, column=end_col)
        ila_b_cell.border = border
        ila_b_cell.alignment = Alignment(wrap_text=True, vertical='center')
        if launch_cells_b and ri in launch_cells_b:
            lc = launch_cells_b[ri]
            ila_b_cell.value = lc['text']
            f, fn = _launch_fill(lc['severity'])
            ila_b_cell.fill = f
            ila_b_cell.font = fn

        for si in range(n_splices):
            col = si + 3
            key = (ri, si)
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='center')

            if key in cells:
                cd = cells[key]
                cell.value = cd['text']
                if cd['is_break']:
                    cell.fill = break_fill
                    cell.font = break_font
                elif cd['is_broke']:
                    cell.fill = broke_fill
                    cell.font = broke_font
                elif cd.get('is_bend'):
                    cell.fill = bend_fill
                    cell.font = bend_font
                elif cd.get('is_bfill'):
                    cell.fill = bfill_fill
                    cell.font = bfill_font
                elif cd.get('is_dead_zone'):
                    cell.fill = dz_fill
                    cell.font = dz_font
                elif cd.get('is_gainer'):
                    cell.fill = gainer_fill
                    cell.font = gainer_font
                elif cd.get('is_b_only'):
                    if cd.get('est_bidir_flagged'):
                        cell.fill = bonly_fill2
                        cell.font = bonly_font2
                    else:
                        cell.fill = bonly_fill
                        cell.font = bonly_font
                elif cd.get('is_a_only'):
                    if cd.get('est_bidir_flagged'):
                        cell.fill = aonly_fill2
                        cell.font = aonly_font2
                    else:
                        cell.fill = aonly_fill
                        cell.font = aonly_font
                else:
                    cell.fill = red_fill
                    cell.font = data_font

    # ── Legend sheet ──
    ws_leg = wb.create_sheet("Legend")
    ws_leg.column_dimensions['A'].width = 14
    ws_leg.column_dimensions['B'].width = 65
    legend_items = [
        ("Pink",       "FFC7CE", "000000", "A+B — Bidirectional reburn: both directions confirmed, bidir loss >= threshold. Needs re-splice."),
        ("Red",        "FF4444", "FFFFFF", "Break — 1F reflective event (clean cut, glass-to-air Fresnel reflection). label: 'BREAK'"),
        ("Red (broke)","FF4444", "FFFFFF", "Broke — fiber trace terminates mid-span (crush / stress fracture).  Rendered with the same red fill as a break; label reads 'broke' or 'BREAK' depending on reflective vs non-reflective signature."),
        ("Blue",       "BDD7EE", "1F4E79", "B-fill — B-direction loss used past a break where A-direction is blind. label: '(B-fill)'"),
        ("Gray",       "BFBFBF", "3F3F3F", "Dead zone — fiber broke on A side AND B trace also ends before reaching the A-break. Neither trace could see this splice for this fiber. Broke cell shows 'F# broke@XXk | DZ lo-hi k'; affected columns show 'F# DZ'."),
        ("Lt. Yellow", "FFF2CC", "7F6000", "A-only, est bidir OK — A saw it, no B entry. Estimated bidir (A/2) is below threshold. label: 'F# .xxx(A) ~.xxxbd'"),
        ("Coral",      "FF7043", "FFFFFF", "A-only, est bidir HIGH — A saw it, no B entry. Estimated bidir (A/2) still exceeds threshold. label: 'F# .xxx(A) ⚠.xxxbd'"),
        ("Lavender",   "E8D5F5", "4B0082", "B-only, est bidir OK — B saw it, no A entry. Estimated bidir (B/2) is below threshold. label: 'F# .xxx(B) ~.xxxbd'"),
        ("Purple",     "C084FC", "1A0033", "B-only, est bidir HIGH — B saw it, no A entry. Estimated bidir (B/2) still exceeds threshold. label: 'F# .xxx(B) ⚠.xxxbd'"),
        ("Yellow",     "FFEB3B", "5D4037", "BEND — event ≥ 0.090 dB at a position more than 150 m from the closure center.  Inspect conduit for pinch or tight bend."),
        ("Orange",     "FFA500", "5D2E00", "LAUNCH — fiber has a launch-end issue.  Loss rule: launch_loss >= -0.5 dB (anything weaker than a -0.5 dB gainer flags).  Reflectance rule: refl > -15 dB (damaged / dirty connector).  Plus missing file, empty event table.  Single tier — no WATCH/REVIEW/HIGH split.  Appears in ILA column.  Distinct from pink A+B reburn."),
        ("Mint Green", "A5D6A7", "1B5E20", "FIELD GAINER — mid-span event whose signed loss is in [-0.7, 0] dB (suspicious near-zero / weak-gainer event).  Excludes events within the launch zone or end-of-fiber region.  Overrides the geometric BEND tag in the [-0.7, -0.090] overlap range."),
    ]
    ws_leg.cell(row=1, column=1, value="Color").font = Font(bold=True, size=10)
    ws_leg.cell(row=1, column=2, value="Meaning").font = Font(bold=True, size=10)
    for i, (name, fc, tc, desc) in enumerate(legend_items, 2):
        c = ws_leg.cell(row=i, column=1, value=name)
        c.fill = PatternFill(start_color=fc, end_color=fc, fill_type="solid")
        c.font = Font(bold=True, size=9, color=tc)
        ws_leg.cell(row=i, column=2, value=desc).font = Font(size=9)

    # ── Column widths ──
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 10
    for si in range(n_splices + 1):
        col_letter = openpyxl.utils.get_column_letter(si + 3)
        ws.column_dimensions[col_letter].width = 22

    ws.freeze_panes = 'C4'

    wb.save(output_path)
    print(f"  Saved: {output_path}")


# ═══════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description='Splice QC report with EXFO-style bidirectional event matching.',
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('dir_a', help='A-direction SOR files directory')
    ap.add_argument('dir_b', nargs='?', help='B-direction SOR files directory')
    ap.add_argument('--output', '-o', default='splice_report_exfo.xlsx')
    ap.add_argument('--threshold', type=float, default=REBURN_THRESHOLD,
                    help=f'Flag threshold in dB (default {REBURN_THRESHOLD})')
    ap.add_argument('--ribbon-size', type=int, default=RIBBON_SIZE)
    ap.add_argument('--site-a', default=None,
                    help='A-end site name (auto-detected from directory names if not set)')
    ap.add_argument('--site-b', default=None,
                    help='B-end site name (auto-detected from directory names if not set)')
    ap.add_argument('--span-km', type=float, default=0,
                    help='Span distance in km (0 = auto-detect)')
    args = ap.parse_args()

    # ── Auto-detect site names from directory names ──
    # Directory names like "NEWELM 15 sec" encode the route as site_a + site_b
    # A-direction dir = A→B (e.g., NEWELM = from NEW to ELM)
    # B-direction dir = B→A (e.g., ELMNEW = from ELM to NEW)
    if args.site_a is None or args.site_b is None:
        import re
        a_base = os.path.basename(args.dir_a.rstrip('/'))
        # Extract the alphabetic prefix (e.g., "NEWELM" from "NEWELM 15 sec")
        alpha = re.match(r'([A-Za-z]+)', a_base)
        if alpha:
            route_str = alpha.group(1).upper()
            # Try to split into two 3-letter site codes (e.g., NEWELM → NEW + ELM)
            if len(route_str) >= 6:
                half = len(route_str) // 2
                if args.site_a is None:
                    args.site_a = route_str[:half]
                if args.site_b is None:
                    args.site_b = route_str[half:]
                print(f"  Auto-detected site names: {args.site_a} → {args.site_b}")
        if args.site_a is None:
            args.site_a = 'A'
        if args.site_b is None:
            args.site_b = 'B'

    print("Loading SOR files...")
    fibers_a, fibers_b = load_all(args.dir_a, args.dir_b)
    n_fibers = max(fibers_a.keys()) if fibers_a else 0
    print(f"  A: {len(fibers_a)} fibers   B: {len(fibers_b)} fibers   max fiber #{n_fibers}")

    # ── Pass 0: Normalize events for splice discovery ──
    # Save original events (needed for trace enhancement), normalize copies
    for r in list(fibers_a.values()) + list(fibers_b.values()):
        r['_raw_events'] = r['events']  # save originals
        r['events'] = _normalize_untrimmed_events(r['events'])

    print("Discovering splice closure positions...")
    splice_candidates = discover_splices(fibers_a)
    real_splices, phantom_zones = refine_closure_centers(
        fibers_a, splice_candidates, return_phantoms=True)
    print(f"  Found {len(real_splices)} real splice closures:")
    for i, sp in enumerate(real_splices, 1):
        ref_km = sp.get('position_km_refined', sp['position_km'])
        offset_m = (ref_km - sp['position_km']) * 1000
        print(f"    Splice {i:2d}: {sp['position_km']:8.2f} km  "
              f"(refined {ref_km:7.3f} km, {offset_m:+5.0f} m offset, "
              f"{sp['count']} fibers, spread {sp.get('position_spread_m', 0):.0f} m)")

    # Every bend / damage zone gets its own named column in the grid, sorted
    # with the real splice columns by physical position.  This mirrors the
    # tech's Cle Elum layout exactly (Splice 12, bends@11.51, Splice 11,
    # damage@17.46, bends@18.86, Splice 10, …).
    if phantom_zones:
        print(f"  Adding {len(phantom_zones)} bend/damage column(s) to the grid:")
        for sp in phantom_zones:
            ref = sp.get('position_km_refined', sp['position_km'])
            print(f"    [{sp['column_kind']:<6}] {ref:7.2f} km")
    splices = sorted(
        list(real_splices) + list(phantom_zones),
        key=lambda sp: sp.get('position_km_refined', sp['position_km']),
    )
    # Re-index splice display numbers for the real closures only
    splice_num = 0
    for sp in splices:
        if sp.get('column_kind') == 'splice':
            splice_num += 1
            sp['splice_display_num'] = splice_num

    # ── Launch-issue detection (must run BEFORE events get normalized again) ──
    first_splice_km = splices[0]['position_km'] if splices else None
    print("\nDetecting launch-end issues...")
    launch_issues = detect_launch_issues(fibers_a, fibers_b, first_splice_km)
    high_n   = sum(1 for v in launch_issues.values() if v['severity'] == 'HIGH')
    review_n = sum(1 for v in launch_issues.values() if v['severity'] == 'REVIEW')
    watch_n  = sum(1 for v in launch_issues.values() if v['severity'] == 'WATCH')
    print(f"  {len(launch_issues)} fibers with launch-end issues "
          f"(HIGH={high_n}, REVIEW={review_n}, WATCH={watch_n})")

    # Auto-detect span (preliminary, from normalized events)
    span_km = args.span_km
    if span_km == 0:
        all_ends = sorted([e['dist_km'] for r in fibers_a.values()
                           for e in r['events'] if e['is_end']])
        if all_ends:
            top_quarter = all_ends[int(len(all_ends) * 0.75):]
            span_km = round(np.median(top_quarter), 2)

    # ── Trace-based enhancement: detect breaks and refine span from raw trace ──
    # Restore original events so trace enhancement can detect untrimmed files
    for r in list(fibers_a.values()) + list(fibers_b.values()):
        r['events'] = r.pop('_raw_events')

    n_trace_breaks = 0
    n_trace_enhanced = 0
    has_trace_data = any(r.get('full_trace') is not None for r in fibers_a.values())
    if has_trace_data:
        print(f"\nTrace analysis: detecting breaks and span boundaries from raw trace...")

        # Phase 1: detect noise floors for all fibers to get population baseline
        all_noise_floors = []
        for r in list(fibers_a.values()) + list(fibers_b.values()):
            trace = r.get('full_trace')
            if trace is None:
                continue
            pts = r['full_points']
            acq = r['acq_range']
            ior_val = r.get('ior', 1.4682)
            launch_idx = _detect_launch_from_trace(trace, pts, acq, ior_val)
            nf_idx = _detect_noise_floor_from_trace(trace, launch_idx, pts, acq, ior_val)
            nf_km = _sample_to_km(nf_idx, ior_val, pts, acq)
            launch_km = _sample_to_km(launch_idx, ior_val, pts, acq)
            all_noise_floors.append(nf_km - launch_km)

        if all_noise_floors:
            pop_noise_floor = np.median(sorted(all_noise_floors)[int(len(all_noise_floors)*0.75):])
            print(f"  Population trace noise floor: {pop_noise_floor:.1f} km from launch")
        else:
            pop_noise_floor = span_km

        # Phase 2: enhance events using population noise floor as reference
        for fnum, r in fibers_a.items():
            _enhance_events_with_trace(r, span_km, pop_noise_floor_km=pop_noise_floor)
            if r.get('_trace_breaks'):
                n_trace_breaks += len(r['_trace_breaks'])
            if r.get('_trace_launch_km') is not None:
                n_trace_enhanced += 1
        for fnum, r in fibers_b.items():
            _enhance_events_with_trace(r, span_km, pop_noise_floor_km=pop_noise_floor)
            if r.get('_trace_breaks'):
                n_trace_breaks += len(r['_trace_breaks'])
        print(f"  Enhanced {n_trace_enhanced} A-fibers, {n_trace_breaks} breaks detected from trace")

        # Re-compute span after trace enhancement (use events, not trace)
        if args.span_km == 0:
            all_ends = sorted([e['dist_km'] for r in fibers_a.values()
                               for e in r['events'] if e['is_end']])
            if all_ends:
                top_quarter = all_ends[int(len(all_ends) * 0.75):]
                span_km = round(np.median(top_quarter), 2)

    print(f"  Span: {span_km} km ({span_km * 3280.84:,.0f} ft)")

    print(f"\nPass 1: Analyzing {len(fibers_a)} fibers at {len(splices)} splice positions "
          f"(threshold={args.threshold:.3f} dB)...")
    results = analyze_all(fibers_a, fibers_b, splices, args.threshold)
    n_p1_bidir  = sum(1 for r in results.values() if r.get('event_source') in ('bidir', 'bidir_grey_b'))
    n_p1_aonly  = sum(1 for r in results.values() if r.get('is_a_only'))
    n_p1_broke  = sum(1 for r in results.values() if r['is_broke'])
    n_p1_break  = sum(1 for r in results.values() if r['is_break'])
    n_p1_bend   = sum(1 for r in results.values() if r.get('is_bend'))
    n_p1_bfill  = sum(1 for r in results.values() if r.get('is_bfill'))
    print(f"  Pass 1 results: {len(results)} events")
    print(f"    A+B bidir:  {n_p1_bidir}")
    print(f"    A-only:     {n_p1_aonly}")
    print(f"    Breaks:     {n_p1_break}")
    print(f"    Broke:      {n_p1_broke}")
    print(f"    Bends:      {n_p1_bend}")
    print(f"    B-fill:     {n_p1_bfill}")

    # APRIL 23 revision: replace the old Pass-2 B-scan with two narrower passes
    #   Pass 2a — A-first standalone event classifier (bends / breaks
    #             at non-closure positions)
    #   Pass 2b — past-break B-fill scan (only uses B direction past
    #             an A-side break)
    print(f"\nPass 2a: Scanning A-direction standalone events (bends / breaks)...")
    a_standalone = scan_a_standalone_events(
        fibers_a, splices, results, span_km,
    )
    n_p2a_bend  = sum(1 for r in a_standalone.values() if r.get('is_bend'))
    n_p2a_break = sum(1 for r in a_standalone.values() if r.get('is_break'))
    print(f"  Pass 2a results: {len(a_standalone)} events "
          f"(bends={n_p2a_bend}, breaks={n_p2a_break})")

    print(f"\nPass 2b: Scanning B-direction PAST A-side breaks (B-fill only)...")
    b_pastbreak = scan_b_past_breaks(
        fibers_a, fibers_b, splices, args.threshold, results, span_km,
    )
    print(f"  Pass 2b results: {len(b_pastbreak)} B-fill events")

    # Merge — Pass 1 takes priority; then standalone; then B-fill
    all_results = {**results, **a_standalone, **b_pastbreak}
    b_results = {**a_standalone, **b_pastbreak}  # kept for any downstream count code

    # Field-gainer annotation — flag mid-span events whose signed loss
    # falls in [-0.7, 0] dB (suspicious near-zero / weak-gainer events).
    n_field_gainers = apply_field_gainer_rule(all_results, span_km)
    print(f"  Field gainers: {n_field_gainers} (loss in "
          f"[{FIELD_GAINER_MIN_DB}, {FIELD_GAINER_MAX_DB}] dB, mid-span)")

    n_total   = len(all_results)
    n_bend    = sum(1 for r in all_results.values() if r.get('is_bend'))
    n_bidir   = sum(1 for r in all_results.values()
                    if r.get('event_source') in ('bidir', 'bidir_grey_a', 'bidir_grey_b'))
    n_a_only  = sum(1 for r in all_results.values() if r.get('is_a_only'))
    n_b_only  = sum(1 for r in all_results.values() if r.get('is_b_only'))
    n_breaks  = sum(1 for r in all_results.values() if r['is_break'])
    n_broke   = sum(1 for r in all_results.values() if r['is_broke'])
    n_bfill   = sum(1 for r in all_results.values() if r.get('is_bfill'))
    n_dz      = sum(1 for r in all_results.values() if r.get('is_dead_zone'))
    n_dz_fibers = len({r['fiber'] for r in all_results.values()
                       if r['is_broke'] and r.get('dead_zone_km')})
    n_reburn  = n_bidir - n_breaks

    # Bend severity breakdown
    n_bend_high   = sum(1 for r in all_results.values()
                        if r.get('is_bend') and r.get('bend_severity') == 'HIGH')
    n_bend_review = sum(1 for r in all_results.values()
                        if r.get('is_bend') and r.get('bend_severity') == 'REVIEW')
    n_bend_watch  = sum(1 for r in all_results.values()
                        if r.get('is_bend') and r.get('bend_severity') == 'WATCH')

    print(f"\nBuilding ribbon grid...")
    cells, launch_cells_a, launch_cells_b = build_ribbon_data(
        all_results, n_fibers, args.ribbon_size, len(splices),
        launch_issues=launch_issues,
    )
    print(f"  {len(cells)} cells with flagged events, "
          f"{len(launch_cells_a)} ribbons with A-launch issues, "
          f"{len(launch_cells_b)} ribbons with B-launch issues")

    print(f"Writing Excel report...")
    write_xlsx(cells, splices, n_fibers, args.ribbon_size, args.output,
               args.site_a, args.site_b, span_km,
               launch_cells_a=launch_cells_a, launch_cells_b=launch_cells_b)

    print(f"\n{'═'*60}")
    print(f"  SPLICE REPORT (EXFO-MATCH + BENDS) COMPLETE")
    print(f"{'═'*60}")
    print(f"  Fibers:       {n_fibers}")
    print(f"  Splices:      {len(splices)}")
    print(f"  Span:         {span_km} km")
    print(f"  Threshold:    {args.threshold:.3f} dB   (bend threshold {BEND_THRESHOLD:.3f} dB, offset > {CLOSURE_MATCH_KM*1000:.0f} m)")
    print(f"  ──────────────────────────────────")
    print(f"  A+B reburns:  {n_reburn}  (pink)   — both directions, bidir >= threshold, near closure center")
    print(f"  Breaks:       {n_breaks}  (red)    — 1F reflective event")
    print(f"  Broke:        {n_broke}  (red)    — trace terminates mid-span (same red fill as break)")
    print(f"  B-fill:       {n_bfill}  (blue)   — B-direction past a break")
    print(f"  Dead zone:    {n_dz}  (gray)   — neither trace could see; {n_dz_fibers} broken fibers have a dead zone")
    print(f"  A-only:       {n_a_only}  (yellow) — A saw it, B did not")
    print(f"  B-only:       {n_b_only}  (purple) — B saw it, A did not  ← EXFO extra")
    print(f"  Bends:        {n_bend}  (yellow) — event >= {BEND_THRESHOLD:.3f} dB, > 150 m from closure center")
    n_gainer = sum(1 for r in all_results.values() if r.get('is_gainer'))
    print(f"  Field gainers:{n_gainer}  (mint)   — mid-span loss in [{FIELD_GAINER_MIN_DB}, {FIELD_GAINER_MAX_DB}] dB")
    print(f"  Launch:       {len(launch_issues)}  (orange) — launch-end issues (single tier)")
    print(f"  ──────────────────────────────────")
    print(f"  Total:        {n_total}")
    print(f"  Output:       {args.output}")
    print()


if __name__ == '__main__':
    main()
