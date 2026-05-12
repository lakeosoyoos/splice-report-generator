"""
make_zach_explanation.py
========================

Per-cell explanation report for Zach.  Walks every flagged event in
the pipeline output and explains, in plain English, why the script
classified it the way it did.

Usage from app.py:

    from make_zach_explanation import build_explanation_pdf
    build_explanation_pdf(
        all_results, splices, launch_issues, span_km,
        site_a, site_b, output_path)

No tech-report comparison.  No trace overlays.  Just a clean
narrative + tabular per-cell breakdown for each category.
"""
from __future__ import annotations
import math

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors as rlcolors
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
)


BLUE = "#1F4E79"
RED  = "#E8461E"
RIBBON_SIZE = 12

# ── colors used by the pipeline ─────────────────────────────────────
COLOR = {
    'pink':         '#FFC7CE',
    'red':          '#FF4444',
    'deep_orange':  '#E64A19',
    'blue':         '#BDD7EE',
    'gray':         '#BFBFBF',
    'yellow':       '#FFEB3B',
    'mint':         '#A5D6A7',
    'lt_yellow':    '#FFF2CC',
    'coral':        '#FF7043',
    'lavender':     '#E8D5F5',
    'purple':       '#C084FC',
    'orange':       '#FFA500',
}


# ─────────────────────────────────────────────────────────────────────
# Per-category explanation builders
# ─────────────────────────────────────────────────────────────────────

def _ribbon_for(fnum, ribbon_size=RIBBON_SIZE):
    return (fnum - 1) // ribbon_size + 1


def _fmt_loss(v):
    if v is None: return '—'
    try: return f"{float(v):+.3f}"
    except Exception: return str(v)


def _fmt_km(v):
    if v is None: return '—'
    try: return f"{float(v):.3f}"
    except Exception: return str(v)


def _b_source_label(r):
    """Friendly B-source label per result dict."""
    if r.get('_b_source') == 'grey_lsa' or r.get('_b_is_grey'):
        return 'grey-LSA (no B event)'
    if r.get('_a_source') == 'grey_lsa' or r.get('_a_is_grey'):
        return 'grey-LSA on A side'
    if r.get('event_source') == 'bidir_grey_b':
        return 'grey-LSA (no B event)'
    if r.get('event_source') == 'bidir_grey_a':
        return 'grey-LSA on A side'
    if r.get('event_source') == 'bfill':
        return 'B event (past A-break)'
    if r.get('b_loss') is not None:
        return 'discrete B event'
    if r.get('a_loss') is not None and r.get('b_loss') is None:
        return 'A-only — no B'
    return '—'


def _explain_pink(r):
    bd = r.get('bidir_loss')
    a = r.get('a_loss')
    b = r.get('b_loss')
    bsrc = _b_source_label(r)
    if bd is None:
        return ("A+B reburn category, but bidir average isn't recorded — "
                "see the per-direction values to understand the source.")
    parts = [
        f"Both directions confirm a real splice loss here.  "
        f"A reads {_fmt_loss(a)} dB, B reads {_fmt_loss(b)} dB ({bsrc}).  "
        f"Bidirectional average = {bd:+.3f} dB ≥ 0.160 dB reburn threshold."
    ]
    if r.get('_b_is_grey') or r.get('_b_source') == 'grey_lsa':
        parts.append("B value came from a wide-LSA grey reading on the "
                     "raw B trace because no discrete B event existed at "
                     "the mirrored position — same technique EXFO "
                     "FastReporter uses when it shows a grey-shaded loss.")
    return ' '.join(parts)


def _explain_bend(r):
    bd = r.get('bidir_loss')
    a = r.get('a_loss')
    b = r.get('b_loss')
    off = r.get('closure_offset_m')
    bsrc = _b_source_label(r)
    pieces = [
        f"Bend classifier fired.  A reads {_fmt_loss(a)} dB, "
        f"B reads {_fmt_loss(b)} dB ({bsrc}), bidir avg = {_fmt_loss(bd)} dB "
        f"(≥ 0.090 dB bend threshold)."
    ]
    if off is not None:
        pieces.append(
            f"The event sits {off:+.0f} m from this fiber's predicted "
            "splice km (per-fiber linear length-model residual)."
        )
    pieces.append(
        "Per-fiber length-model residual ≥ 150 m AND a narrow-window LSA "
        "on the raw A trace at the predicted splice km confirmed a "
        "separate loss step (≥ 0.030 dB).  That's the bend signature: a "
        "real splice exists where the model predicts AND a separate event "
        "exists at the candidate position."
    )
    return ' '.join(pieces)


def _explain_ref(r):
    bd = r.get('bidir_loss')
    refl = r.get('fresnel')
    src = r.get('event_source', '')
    if src == 'ref_bidir_ghost':
        return (
            f"Bidirectional ghost reflection.  Mid-span 1F reflective "
            f"event with near-zero loss ({_fmt_loss(bd)} dB) and a faint "
            f"Fresnel reflection ({(refl or 0):.0f} dB) — but the SAME "
            "feature shows in BOTH directions at the mirror-matched km "
            "(±100 m), so it is NOT instrument noise.  Real physical "
            "cause: faint connector pair, mechanical splice, angled "
            "cleave, or a downstream reflector creating a ghost.  "
            "Would slip past the regular loss-threshold gates because "
            "the loss is below 0.030 dB; the bidirectional-mirror check "
            "is what lets us call it out."
        )
    return (
        f"Reflective event (1F type) with Fresnel reflection "
        f"{(refl or 0):.0f} dB and bidir loss {_fmt_loss(bd)} dB.  "
        "The fiber's trace continues past this point with real events "
        "downstream and the EOF still 3+ km farther — so this is NOT a "
        "break.  It's classified as an in-line reflective event "
        "(connector pair, mechanical splice, angled cleave).  Cell label "
        "carries both the loss and the reflection magnitude."
    )


def _explain_break(r):
    bd = r.get('bidir_loss')
    refl = r.get('fresnel')
    return (
        f"Reflective event (1F type) with Fresnel reflection "
        f"{(refl or 0):.0f} dB and loss {_fmt_loss(bd)} dB.  The fiber's "
        "trace TERMINATES near this point (no real events downstream, "
        "or EOF too close).  Classified as a real BREAK — physical fault "
        "at this position."
    )


def _explain_broke(r):
    label = r.get('label') or ''
    return (
        f"Fiber's A trace terminates mid-span — the OTDR sees the trace "
        "die before reaching the far end.  Cell label: '" + label + "'.  "
        "Past-A-break B-fill (Pass 2b) walks downstream splices to "
        "recover what's still visible from the B end."
    )


def _explain_bfill(r):
    a = r.get('a_loss'); b = r.get('b_loss'); bd = r.get('bidir_loss')
    return (
        f"A side is blind here (fiber broken upstream).  B trace shows a "
        f"discrete event at this closure with loss {_fmt_loss(b)} dB.  "
        "Past-A-break B-fill recovers it into the report — without this "
        "step, the cell would be empty for every closure past the break."
    )


def _explain_gainer(r):
    a = r.get('a_loss'); b = r.get('b_loss'); bd = r.get('bidir_loss')
    return (
        "Strict bidirectional gainer signature: "
        f"A = {_fmt_loss(a)} dB, B = {_fmt_loss(b)} dB, bidir avg = "
        f"{_fmt_loss(bd)} dB (in [−0.7, 0] dB gainer range).  Both A and "
        "B are real event measurements (no grey-LSA), and the two "
        "directions have opposite signs — the canonical gainer "
        "fingerprint (one direction sees apparent gain, the other sees "
        "matching loss)."
    )


def _explain_a_only(r):
    a = r.get('a_loss')
    est = r.get('est_bidir')
    flagged_high = r.get('est_bidir_flagged')
    note = ("Estimated bidir (A/2) still clears 0.160 dB, so this cell "
            "is rendered in coral as a HIGH-priority A-only — worth a "
            "manual look in EXFO.") if flagged_high else (
            "Estimated bidir (A/2) is below 0.160 dB so the cell is "
            "rendered in light yellow — informational, not a reburn-grade "
            "alarm.")
    return (
        f"A direction reads {_fmt_loss(a)} dB.  B has no event at this "
        f"position and the wide-LSA grey reading didn't pull a usable "
        f"value either.  Estimated bidirectional = {_fmt_loss(est)} dB "
        "(half the A-side magnitude as a rough proxy).  " + note
    )


def _explain_b_only(r):
    b = r.get('b_loss')
    est = r.get('est_bidir')
    flagged_high = r.get('est_bidir_flagged')
    note = ("Estimated bidir (B/2) still clears 0.160 dB, so this cell "
            "is rendered in purple — HIGH-priority B-only.") if flagged_high else (
            "Estimated bidir (B/2) is below 0.160 dB so the cell is "
            "rendered in lavender — informational, not a reburn-grade "
            "alarm.")
    return (
        f"B direction reads {_fmt_loss(b)} dB.  A has no event at this "
        f"position.  Estimated bidir = {_fmt_loss(est)} dB.  " + note
    )


def _explain_dead_zone(r):
    return (
        "Both A and B traces are blind here.  Fiber is broken on the A "
        "side AND the B trace also stops short of reaching the A-break "
        "position — neither direction can observe this splice for this "
        "fiber.  Cell renders gray with a 'DZ' label.  No data; no flag "
        "magnitude available."
    )


def _explain_launch(info):
    sev = info.get('severity', '')
    a_tags = info.get('a_tags', []) or []
    b_tags = info.get('b_tags', []) or []
    pieces = [f"Launch-end issue, severity {sev}."]
    if a_tags:
        pieces.append("A-side tags: " + ', '.join(a_tags) + ".")
    if b_tags:
        pieces.append("B-side tags: " + ', '.join(b_tags) + ".")
    pieces.append(
        "Triggered by one or more of: missing event table, fiber ends "
        "within 2 km of launch, launch-connector loss exceeds 1 dB, "
        "launch-connector reflectance at or above −49.9 dB, "
        "tailbox-connector reflectance at or above −49.9 dB / missing "
        "tailbox (bare-glass cable end), OR the fiber was shot with a "
        "different acquisition duration / pulse width than the majority "
        "of fibers in this direction (DURATION_MISMATCH — FQA fails a "
        "span when traces weren't all shot the same length).  Healthy "
        "buried launch and tailbox both reflect −50 to −55 dB; values "
        "closer to zero indicate damaged / dirty / partially-cut "
        "connector — or, for BAD_TAILBOX_REFL, that the cable end has "
        "no tailbox connector installed at all.")
    return ' '.join(pieces)


# ─────────────────────────────────────────────────────────────────────
# Table builder
# ─────────────────────────────────────────────────────────────────────

def _make_table(headers, rows, header_fill):
    """Build a styled reportlab Table."""
    data = [headers] + rows
    tbl = Table(data, repeatRows=1)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), rlcolors.HexColor(header_fill)),
        ('TEXTCOLOR',  (0, 0), (-1, 0), rlcolors.white),
        ('FONTNAME',   (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0, 0), (-1, -1), 7),
        ('ALIGN',      (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN',     (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING',    (0, 0), (-1, -1), 3),
        ('LEFTPADDING',   (0, 0), (-1, -1), 4),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
        ('GRID',       (0, 0), (-1, -1), 0.25, rlcolors.HexColor('#888888')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1),
            [rlcolors.white, rlcolors.HexColor('#F5F5F5')]),
    ])
    tbl.setStyle(style)
    return tbl


def _row_for(r, splices, ribbon_size, expl_func):
    si = r.get('splice_idx')
    fnum = r.get('fiber')
    sp = splices[si] if si is not None and si < len(splices) else None
    sp_km = (sp.get('position_km_display',
                    sp.get('position_km_refined', sp['position_km']))
              if sp else None)
    ribbon = _ribbon_for(fnum, ribbon_size)
    return [
        f"S{si+1}" if si is not None else '—',
        _fmt_km(sp_km),
        f"R{ribbon}",
        f"F{fnum}",
        _fmt_loss(r.get('a_loss')),
        _fmt_loss(r.get('b_loss')),
        _fmt_loss(r.get('bidir_loss')),
        _b_source_label(r),
        Paragraph(expl_func(r), getSampleStyleSheet()['BodyText']),
    ]


# ─────────────────────────────────────────────────────────────────────
# Main entrypoint
# ─────────────────────────────────────────────────────────────────────

def build_explanation_pdf(all_results, splices, launch_issues, span_km,
                          site_a, site_b, output_path,
                          ribbon_size=RIBBON_SIZE,
                          reburn_threshold=0.160,
                          bend_threshold=0.090):
    """Render the full explanation PDF to `output_path`."""
    doc = SimpleDocTemplate(
        output_path, pagesize=letter,
        leftMargin=0.45*inch, rightMargin=0.45*inch,
        topMargin=0.5*inch, bottomMargin=0.5*inch,
        title=f"Explanation for Zach — {site_a} → {site_b}",
        author="Splice Report pipeline",
    )
    ss = getSampleStyleSheet()
    st = {
        'title': ParagraphStyle('Title', parent=ss['Heading1'],
                                 fontSize=20, leading=24,
                                 textColor=rlcolors.HexColor(BLUE),
                                 alignment=TA_LEFT,
                                 spaceAfter=6),
        'h1':    ParagraphStyle('H1', parent=ss['Heading1'],
                                 fontSize=14, leading=18,
                                 textColor=rlcolors.HexColor(BLUE),
                                 spaceBefore=10, spaceAfter=4),
        'h2':    ParagraphStyle('H2', parent=ss['Heading2'],
                                 fontSize=11, leading=14,
                                 textColor=rlcolors.HexColor(RED),
                                 spaceBefore=8, spaceAfter=3),
        'body':  ParagraphStyle('Body', parent=ss['BodyText'],
                                 fontSize=9, leading=12,
                                 spaceAfter=4),
        'lede':  ParagraphStyle('Lede', parent=ss['BodyText'],
                                 fontSize=10, leading=14,
                                 textColor=rlcolors.HexColor('#222'),
                                 spaceAfter=8),
        'small': ParagraphStyle('Small', parent=ss['BodyText'],
                                 fontSize=8, leading=10),
    }

    story = []

    # ── Bucket every flagged event by category ──────────────────────
    pinks    = []
    bends    = []
    refs     = []
    breaks_  = []
    brokes   = []
    bfills   = []
    gainers  = []
    a_onlys  = []
    b_onlys  = []
    dz       = []
    for k, r in all_results.items():
        if not isinstance(r, dict):
            continue
        # Priority: break > broke > ref > bend > gainer > bfill > a_only > b_only > pink
        if r.get('is_dead_zone'):
            dz.append(r); continue
        if r.get('is_break'):
            breaks_.append(r); continue
        if r.get('is_broke'):
            brokes.append(r); continue
        if r.get('is_ref'):
            refs.append(r); continue
        if r.get('is_bend'):
            bends.append(r); continue
        if r.get('is_gainer'):
            gainers.append(r); continue
        if r.get('is_bfill'):
            bfills.append(r); continue
        if r.get('is_a_only'):
            a_onlys.append(r); continue
        if r.get('is_b_only'):
            b_onlys.append(r); continue
        if r.get('is_flagged'):
            pinks.append(r)

    def _sort(rs):
        return sorted(rs, key=lambda r: (
            r.get('splice_idx') if r.get('splice_idx') is not None else 999,
            r.get('fiber') or 0))

    pinks   = _sort(pinks)
    bends   = _sort(bends)
    refs    = _sort(refs)
    breaks_ = _sort(breaks_)
    brokes  = _sort(brokes)
    bfills  = _sort(bfills)
    gainers = _sort(gainers)
    a_onlys = _sort(a_onlys)
    b_onlys = _sort(b_onlys)
    dz      = _sort(dz)

    # ── Title page ─────────────────────────────────────────────────
    story.append(Paragraph(f"Explanation Report — {site_a} → {site_b}",
                            st['title']))
    story.append(Paragraph(
        "Per-cell explanation of every flagged event the script "
        "produced.  No comparison to any tech report — this report "
        "stands alone, walks every category, and explains why each "
        "fiber landed in that category.  Read top-to-bottom or jump "
        "to a section by category.",
        st['lede']))

    counts = [
        ("Pink — A+B reburn",      len(pinks),   COLOR['pink']),
        ("Yellow — Bend",          len(bends),   COLOR['yellow']),
        ("Deep orange — Ref",      len(refs),    COLOR['deep_orange']),
        ("Red — Break",            len(breaks_), COLOR['red']),
        ("Red — Broke",            len(brokes),  COLOR['red']),
        ("Blue — B-fill",          len(bfills),  COLOR['blue']),
        ("Mint — Field gainer",    len(gainers), COLOR['mint']),
        ("Lt yellow / coral — A-only", len(a_onlys), COLOR['coral']),
        ("Lavender / purple — B-only", len(b_onlys), COLOR['purple']),
        ("Gray — Dead zone",       len(dz),      COLOR['gray']),
        ("Orange — Launch issue",  len(launch_issues or {}), COLOR['orange']),
    ]
    summary_rows = [[
        Paragraph(f"<font color='{c[2]}'>■</font> &nbsp;{c[0]}",
                   st['body']),
        f"{c[1]}"
    ] for c in counts]
    summary_tbl = Table([['Category', '#']] + summary_rows,
                        colWidths=[3.5*inch, 0.7*inch])
    summary_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), rlcolors.HexColor(BLUE)),
        ('TEXTCOLOR',  (0, 0), (-1, 0), rlcolors.white),
        ('FONTNAME',   (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0, 0), (-1, -1), 9),
        ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID',       (0, 0), (-1, -1), 0.25, rlcolors.HexColor('#888888')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1),
            [rlcolors.white, rlcolors.HexColor('#F5F5F5')]),
    ]))
    story.append(summary_tbl)
    story.append(Spacer(1, 12))
    story.append(Paragraph(
        f"<b>Run parameters.</b>  Span ≈ {span_km:.2f} km, "
        f"{len(splices)} splice columns, ribbon size = {ribbon_size}, "
        f"reburn threshold = {reburn_threshold:.3f} dB, bend threshold = "
        f"{bend_threshold:.3f} dB.",
        st['body']))
    story.append(PageBreak())

    # ── Section helper ─────────────────────────────────────────────
    def _emit_section(num, title, narrative, header_fill, rows, expl_func,
                       headers=None):
        story.append(Paragraph(f"{num}. {title}  ({len(rows)} cells)",
                                st['h1']))
        story.append(Paragraph(narrative, st['body']))
        if not rows:
            story.append(Paragraph(
                "<i>None on this cable.</i>", st['body']))
            return
        hdrs = headers or ['Splice', 'Cable km', 'Ribbon', 'Fiber',
                           'A loss', 'B loss', 'Bidir', 'B src',
                           'Why flagged']
        body_rows = [_row_for(r, splices, ribbon_size, expl_func)
                     for r in rows]
        col_widths = [0.45*inch, 0.6*inch, 0.45*inch, 0.45*inch,
                      0.55*inch, 0.55*inch, 0.55*inch,
                      1.1*inch, 3.1*inch]
        tbl = Table([hdrs] + body_rows, colWidths=col_widths,
                     repeatRows=1)
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), rlcolors.HexColor(header_fill)),
            ('TEXTCOLOR',  (0, 0), (-1, 0), rlcolors.HexColor('#222222')),
            ('FONTNAME',   (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0, 0), (-1, -1), 7),
            ('ALIGN',      (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN',     (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
            ('GRID',       (0, 0), (-1, -1), 0.25, rlcolors.HexColor('#888888')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1),
                [rlcolors.white, rlcolors.HexColor('#F5F5F5')]),
        ]))
        story.append(tbl)
        story.append(PageBreak())

    # ── Sections ───────────────────────────────────────────────────
    _emit_section(
        "1", "Pink — A+B Bidirectional Reburns",
        "Both directions of the OTDR confirm a real splice loss at "
        "this closure.  Bidirectional average ≥ 0.160 dB threshold.  "
        "These are the canonical reburn candidates — clean fusion "
        "splices with too much loss for the spec.  Each row below "
        "shows the per-direction values, the B source (whether B came "
        "from a discrete event or a wide-LSA grey reading on the raw "
        "trace), and a one-line reason.",
        COLOR['pink'], pinks, _explain_pink)

    _emit_section(
        "2", "Yellow — Bends",
        "Two-test classifier (April 28).  Test 1 fits a per-fiber "
        "linear length model through every other closure's event; if "
        "the candidate event lies more than 150 m off the model's "
        "prediction, it's a bend candidate.  Test 2 confirms by "
        "reading a narrow-window LSA on the raw A trace at the "
        "predicted splice km — a real bend means a separate splice "
        "exists at the predicted km AND a separate event at the "
        "candidate position.  Loss must also be positive (≥ 0.090 dB).",
        COLOR['yellow'], bends, _explain_bend)

    _emit_section(
        "3", "Deep Orange — In-line Reflective Events",
        "Reflective events with strong Fresnel reflection where the "
        "fiber's trace clearly continues past the event.  These are "
        "connector pairs, mechanical splices, angled cleaves — any "
        "optical interface that produces a Fresnel spike but passes "
        "light through.  Distinct from BREAK red because the fiber "
        "isn't broken.  Cell label format: 'F# ref .xxx (refl -XX dB)'.",
        COLOR['deep_orange'], refs, _explain_ref)

    _emit_section(
        "4", "Red — Break",
        "Reflective events with strong Fresnel reflection where the "
        "fiber's trace TERMINATES near the event (no real events "
        "downstream, or EOF too close).  These are physical faults — "
        "cuts, hard breaks, severed fiber.  Distinct from in-line "
        "reflective events (Section 3) where the trace continues.",
        COLOR['red'], breaks_, _explain_break)

    _emit_section(
        "4b", "Red — Broke",
        "Fiber's A trace terminates mid-span — the OTDR sees the "
        "trace die before reaching the far end.  Past-A-break B-fill "
        "(Section 5) walks downstream splices to recover what's still "
        "visible from the B end.",
        COLOR['red'], brokes, _explain_broke)

    _emit_section(
        "5", "Blue — B-fill (past A-side break)",
        "When a fiber is broken on the A side, the A direction can't "
        "see anything past the break.  The B direction shoots from the "
        "opposite end of the cable, so it can still measure the "
        "section between the A-break and the B-launch.  Pass 2b walks "
        "every B event in that 'past-the-A-break' zone, mirrors the "
        "position into A-frame coordinates, finds the nearest splice "
        "column, and writes a blue B-fill cell.",
        COLOR['blue'], bfills, _explain_bfill)

    _emit_section(
        "6", "Mint — Field Gainers",
        "Strict bidirectional rule (April 28).  Bidir avg lies in "
        "[−0.7, 0] dB AND BOTH a_loss and b_loss are real event "
        "measurements (no grey-LSA on either side) AND the two "
        "directions have OPPOSITE signs — the canonical gainer "
        "fingerprint (one direction sees apparent gain because the "
        "scattering coefficient rises, the other sees matching loss "
        "going the opposite way).",
        COLOR['mint'], gainers, _explain_gainer)

    _emit_section(
        "7", "Light yellow / Coral — A-only",
        "A direction has a discrete event but B has nothing at this "
        "position (no event AND no usable wide-LSA grey).  Cell shows "
        "A's loss with a (A) marker.  Light yellow when the estimated "
        "bidir (A/2) is below 0.160 dB.  Coral when A/2 still clears "
        "the threshold — that's a HIGH-priority A-only worth a manual "
        "look in EXFO.",
        COLOR['lt_yellow'], a_onlys, _explain_a_only)

    _emit_section(
        "8", "Lavender / Purple — B-only",
        "B direction has a discrete event but A has nothing at this "
        "position.  Cell shows B's loss with a (B) marker.  Lavender "
        "when the estimated bidir (B/2) is below 0.160 dB; purple "
        "when B/2 still clears the threshold (HIGH-priority B-only).",
        COLOR['lavender'], b_onlys, _explain_b_only)

    _emit_section(
        "9", "Gray — Dead Zone",
        "Both A and B traces are blind here.  The fiber is broken on "
        "the A side AND the B trace also stops short of reaching the "
        "A-break position — neither direction can observe these "
        "splices for this fiber.  Marked gray with a 'DZ' label.",
        COLOR['gray'], dz, _explain_dead_zone)

    # Launch section uses a different shape — index by fiber, not cell.
    story.append(Paragraph(
        f"10. Orange — Launch-end Issues  ({len(launch_issues or {})} fibers)",
        st['h1']))
    story.append(Paragraph(
        "Fibers with problems at the launch-end connector OR the "
        "tailbox-end connector — broken at the launch, damaged or dirty "
        "connector, missing event table, missing tailbox (bare-glass "
        "cable end), or reflectance outside the healthy range.  "
        "Rule: <b>NO_EVENTS</b>, <b>IMMEDIATE_END</b> (fiber ends "
        "within 2 km of launch), <b>HIGH_LAUNCH_LOSS</b> (launch "
        "connector loss > 1 dB), <b>BAD_LAUNCH_REFL</b> (launch "
        "reflectance at or above −49.9 dB), <b>BAD_TAILBOX_REFL</b> "
        "(tailbox connector reflectance at or above −49.9 dB, OR no "
        "tailbox event at all with a bad reflection on the 1E end — "
        "indicates a missing / dirty tailbox), <b>DURATION_MISMATCH</b> "
        "(this fiber's acquisition duration signature — N_averages, "
        "pulse width — differs from the majority of fibers shot in this "
        "direction; FQA fails a span when traces weren't all shot the "
        "same length).  Healthy buried launch and tailbox both fall in "
        "the −50 to −55 dB range.  Renders in the dedicated ILA:A / "
        "ILA:B columns of the xlsx.",
        st['body']))
    if launch_issues:
        rows = []
        for fnum in sorted(launch_issues.keys()):
            info = launch_issues[fnum]
            rows.append([
                f"F{fnum}",
                f"R{_ribbon_for(fnum, ribbon_size)}",
                info.get('severity', '') or '—',
                ', '.join(info.get('a_tags') or []) or '—',
                ', '.join(info.get('b_tags') or []) or '—',
                Paragraph(_explain_launch(info),
                           getSampleStyleSheet()['BodyText']),
            ])
        tbl = Table([['Fiber', 'Ribbon', 'Severity',
                      'A tags', 'B tags', 'Why flagged']] + rows,
                     colWidths=[0.5*inch, 0.5*inch, 0.7*inch,
                                1.2*inch, 1.2*inch, 3.6*inch],
                     repeatRows=1)
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), rlcolors.HexColor(COLOR['orange'])),
            ('FONTNAME',   (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0, 0), (-1, -1), 7),
            ('ALIGN',      (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN',     (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('GRID',       (0, 0), (-1, -1), 0.25, rlcolors.HexColor('#888888')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1),
                [rlcolors.white, rlcolors.HexColor('#F5F5F5')]),
        ]))
        story.append(tbl)
    else:
        story.append(Paragraph("<i>No launch-end issues on this cable.</i>",
                                st['body']))

    doc.build(story)
    return output_path
