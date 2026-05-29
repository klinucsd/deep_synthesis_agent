"""
comprehensive_report — render a per-submission PDF report combining the
Science + Practitioner rubric reviews and the boss's per-submission
synthesis questions.

Pair module: skills/wildfire-comprehensive-report/SKILL.md.

The SKILL.md instructs the agent to call `generate_report(ws_root)`.
This module:
  1. Loads both _rubric_science.json and _rubric_practitioner.json
     (raises if either is missing — both are required deliverables).
  2. Loads the synthesis_questions.json mapping.
  3. Builds a reportlab PDF with cover page, executive summary,
     full Science rubric section, full Practitioner rubric section,
     synthesis-question section (per-submission scope only), and a
     compact file-inventory appendix.
  4. Writes the PDF next to the JSONs in the workspace folder.

No LLM calls happen here. The PDF is composed purely from the JSONs
on disk; the agent's job upstream is just to call `generate_report()`.
"""

import json
import re
from datetime import datetime, timezone
from pathlib import Path

# reportlab — pre-installed in Sage v1.1.27+
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
    Table, TableStyle, PageBreak, KeepTogether,
)
from reportlab.pdfgen import canvas as rl_canvas

# Import the rubric schemas from the wildfire-rubric-science skill
import sys
_THIS_DIR = Path(__file__).resolve().parent
_SCIENCE_DIR = _THIS_DIR.parent / "wildfire-rubric-science"
if str(_SCIENCE_DIR) not in sys.path:
    sys.path.insert(0, str(_SCIENCE_DIR))
from wildfire_rubric import SCIENCE_RUBRIC, PRACTITIONER_RUBRIC


# ---------------------------------------------------------------------------
# Design tokens — restrained, professional palette
# ---------------------------------------------------------------------------

COLOR_PRIMARY      = colors.HexColor("#1B2A4E")   # deep navy
COLOR_SECONDARY    = colors.HexColor("#475569")   # slate
COLOR_MUTED        = colors.HexColor("#64748B")   # warm gray
COLOR_LIGHT_BG     = colors.HexColor("#F8FAFC")   # very light gray-blue
COLOR_BORDER       = colors.HexColor("#CBD5E1")   # subtle border
COLOR_ANSWERED     = colors.HexColor("#0F766E")   # teal — "answered" markers
COLOR_BRANCHING    = colors.HexColor("#92400E")   # amber — "n/a" markers
COLOR_GAP          = colors.HexColor("#9F1239")   # rose — "evidence gap" markers
COLOR_HIGHLIGHT_BG = colors.HexColor("#EFF6FF")   # very soft blue for answer callouts

FONT_BODY      = "Helvetica"
FONT_BODY_B    = "Helvetica-Bold"
FONT_BODY_I    = "Helvetica-Oblique"
FONT_BODY_BI   = "Helvetica-BoldOblique"

PAGE_W, PAGE_H = letter
MARGIN_X = 0.85 * inch
MARGIN_TOP = 0.9 * inch
MARGIN_BOTTOM = 0.85 * inch


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

def _styles():
    """Build a ParagraphStyle bundle for the report."""
    base = getSampleStyleSheet()
    s = {}

    # Cover page
    s["cover_title"] = ParagraphStyle(
        "cover_title", parent=base["Title"],
        fontName=FONT_BODY_B, fontSize=30, leading=36,
        textColor=COLOR_PRIMARY, alignment=TA_CENTER, spaceAfter=12,
    )
    s["cover_subtitle"] = ParagraphStyle(
        "cover_subtitle", parent=base["Normal"],
        fontName=FONT_BODY, fontSize=14, leading=20,
        textColor=COLOR_SECONDARY, alignment=TA_CENTER, spaceAfter=8,
    )
    s["cover_meta"] = ParagraphStyle(
        "cover_meta", parent=base["Normal"],
        fontName=FONT_BODY, fontSize=11, leading=16,
        textColor=COLOR_MUTED, alignment=TA_CENTER, spaceAfter=4,
    )
    s["cover_badge"] = ParagraphStyle(
        "cover_badge", parent=base["Normal"],
        fontName=FONT_BODY_B, fontSize=11, leading=14,
        textColor=COLOR_PRIMARY, alignment=TA_CENTER,
    )

    # Section headers
    s["h1"] = ParagraphStyle(
        "h1", parent=base["Heading1"],
        fontName=FONT_BODY_B, fontSize=20, leading=26,
        textColor=COLOR_PRIMARY, spaceBefore=0, spaceAfter=12,
    )
    s["h2"] = ParagraphStyle(
        "h2", parent=base["Heading2"],
        fontName=FONT_BODY_B, fontSize=14, leading=18,
        textColor=COLOR_PRIMARY, spaceBefore=14, spaceAfter=6,
    )
    s["h3"] = ParagraphStyle(
        "h3", parent=base["Heading3"],
        fontName=FONT_BODY_B, fontSize=11, leading=14,
        textColor=COLOR_SECONDARY, spaceBefore=8, spaceAfter=4,
    )

    # Body
    s["body"] = ParagraphStyle(
        "body", parent=base["BodyText"],
        fontName=FONT_BODY, fontSize=10, leading=14,
        textColor=COLOR_PRIMARY, alignment=TA_JUSTIFY, spaceAfter=6,
    )
    s["body_muted"] = ParagraphStyle(
        "body_muted", parent=base["BodyText"],
        fontName=FONT_BODY_I, fontSize=9, leading=12,
        textColor=COLOR_MUTED, alignment=TA_LEFT, spaceAfter=4,
    )
    s["question_text"] = ParagraphStyle(
        "question_text", parent=base["BodyText"],
        fontName=FONT_BODY_B, fontSize=10, leading=14,
        textColor=COLOR_PRIMARY, alignment=TA_LEFT,
        spaceBefore=10, spaceAfter=4,
    )
    s["answer_text"] = ParagraphStyle(
        "answer_text", parent=base["BodyText"],
        fontName=FONT_BODY, fontSize=10, leading=13,
        textColor=COLOR_PRIMARY, alignment=TA_LEFT,
        leftIndent=10, rightIndent=10,
        spaceBefore=0, spaceAfter=4,
        backColor=COLOR_HIGHLIGHT_BG,
        borderColor=COLOR_BORDER, borderWidth=0,
        borderPadding=6,
    )
    s["answer_marker"] = ParagraphStyle(
        "answer_marker", parent=base["BodyText"],
        fontName=FONT_BODY_B, fontSize=10, leading=13,
        textColor=COLOR_ANSWERED, alignment=TA_LEFT,
        leftIndent=10, spaceBefore=0, spaceAfter=2,
    )
    s["branching_marker"] = ParagraphStyle(
        "branching_marker", parent=base["BodyText"],
        fontName=FONT_BODY_I, fontSize=9, leading=12,
        textColor=COLOR_BRANCHING, alignment=TA_LEFT,
        leftIndent=10, spaceBefore=0, spaceAfter=2,
    )
    s["gap_marker"] = ParagraphStyle(
        "gap_marker", parent=base["BodyText"],
        fontName=FONT_BODY_I, fontSize=9, leading=12,
        textColor=COLOR_GAP, alignment=TA_LEFT,
        leftIndent=10, spaceBefore=0, spaceAfter=2,
    )
    s["notes"] = ParagraphStyle(
        "notes", parent=base["BodyText"],
        fontName=FONT_BODY_I, fontSize=9, leading=12,
        textColor=COLOR_MUTED, alignment=TA_LEFT,
        leftIndent=10, spaceBefore=2, spaceAfter=4,
    )
    s["source_attribution"] = ParagraphStyle(
        "source_attribution", parent=base["BodyText"],
        fontName=FONT_BODY_I, fontSize=8, leading=11,
        textColor=COLOR_MUTED, alignment=TA_LEFT,
        leftIndent=10, spaceBefore=2, spaceAfter=4,
    )

    return s


# ---------------------------------------------------------------------------
# Page template — header/footer on every page after the cover
# ---------------------------------------------------------------------------

def _draw_header_footer(canvas, doc, workspace_name):
    """Draw the header rule + footer workspace name on every content page.
    The right-side "Page X of Y" is drawn by NumberedCanvas at save time —
    NOT here — to avoid overlap from two paint passes."""
    canvas.saveState()
    canvas.setFont(FONT_BODY, 8)
    canvas.setFillColor(COLOR_MUTED)
    canvas.drawString(MARGIN_X, 0.4 * inch, workspace_name)
    # Header: thin colored rule + project name
    canvas.setStrokeColor(COLOR_BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN_X, PAGE_H - MARGIN_TOP + 0.25 * inch,
                PAGE_W - MARGIN_X, PAGE_H - MARGIN_TOP + 0.25 * inch)
    canvas.setFont(FONT_BODY, 8)
    canvas.setFillColor(COLOR_MUTED)
    canvas.drawString(MARGIN_X, PAGE_H - MARGIN_TOP + 0.32 * inch,
                      "Fire Risk Modeling Exercise — Comprehensive Review")
    canvas.restoreState()


def _draw_cover_decoration(canvas, doc):
    """Cover page has no header/footer, just a thin top accent rule."""
    canvas.saveState()
    canvas.setFillColor(COLOR_PRIMARY)
    canvas.rect(0, PAGE_H - 0.5 * inch, PAGE_W, 0.5 * inch, fill=1, stroke=0)
    # Tagline at the very bottom of the cover
    canvas.setFont(FONT_BODY_I, 8)
    canvas.setFillColor(COLOR_MUTED)
    canvas.drawCentredString(PAGE_W / 2, 0.4 * inch,
                             "Generated by an AI review agent")
    canvas.restoreState()


# ---------------------------------------------------------------------------
# Rendering helpers
# ---------------------------------------------------------------------------

_ESCAPE_TABLE = {"&": "&amp;", "<": "&lt;", ">": "&gt;"}

def _escape(text):
    """HTML-escape a string for safe inclusion in reportlab Paragraph()."""
    if text is None:
        return ""
    text = str(text)
    for k, v in _ESCAPE_TABLE.items():
        text = text.replace(k, v)
    return text


def _format_answer(spec, answer):
    """Return a list of (paragraph_style_key, text) tuples representing an
    answered question's value + notes."""
    out = []
    t = spec["type"]

    if t == "single_select":
        v = answer.get("value")
        if v:
            out.append(("answer_marker",
                        f"✓ {_escape(_strip_sage_naming(v))}"))
    elif t == "multi_select":
        sel = answer.get("selected") or []
        if sel:
            checks = " ".join(f"✓ {_escape(_strip_sage_naming(s))}"
                              for s in sel)
            out.append(("answer_marker", checks))
        elif "selected" in answer and not sel:
            out.append(("notes", "(none selected)"))
    elif t == "text":
        v = answer.get("value")
        if v:
            out.append(("answer_marker", _escape(_strip_sage_naming(v))))
    elif t == "narrative":
        # Narrative-only — the substance is in notes; nothing extra to mark.
        pass

    notes = (answer or {}).get("notes")
    if notes:
        out.append(("notes", _escape(_strip_sage_naming(notes))))

    return out


def _coerce_reason(reason):
    """Coerce a branching/gap reason to a clean string.

    Some older rubric JSONs (pre-2026-05-23 validation fix) accidentally
    stored a dict of {char: reason} when the agent called
    `branch_skip(child_ids=<a string>)`. Recover the underlying reason
    rather than rendering the broken dict literally."""
    if isinstance(reason, str):
        return reason
    if isinstance(reason, dict):
        # Recover the common reason if all values are the same string
        vals = list(reason.values())
        if vals and all(isinstance(v, str) and v == vals[0] for v in vals):
            return vals[0]
        return "(skipped — reason data malformed in source JSON)"
    return str(reason)


_SKIP_PREFIXES = ("not applicable — ", "not applicable —",
                  "cascade — ", "cascade —")


def _strip_skip_prefix(text):
    """Remove a leading "not applicable — " or "cascade — " marker so the
    renderer's own prefix doesn't double up. branch_skip/cascade_skip in
    wildfire_rubric.py write reasons that begin with these markers; the
    PDF renderer prepends its own human-readable prefix."""
    if not text:
        return text
    low = text.lstrip().lower()
    for p in _SKIP_PREFIXES:
        if low.startswith(p):
            return text.lstrip()[len(p):].lstrip()
    return text


def _format_skipped(qid, reason_text, kind):
    """Return paragraph tuples for a skipped question (branching or gap)."""
    clean = _strip_skip_prefix(_coerce_reason(reason_text))
    if kind == "branching":
        return [("branching_marker", f"Not applicable — {_escape(clean)}")]
    else:  # gap
        return [("gap_marker", f"Not addressed in available materials — {_escape(clean)}")]


_SAGE_REPLACEMENTS = [
    ("Sage-automated", "automated review"),
    ("Sage automated", "automated review"),
    ("Sage-AutoFill", "automated review"),
    ("Sage AutoFill", "automated review"),
    ("Sage's", "the review agent's"),
    ("Sage", "the review agent"),
]


def _strip_sage_naming(text):
    """Replace Sage product references in agent-written answer text with
    neutral wording. The boss-facing report must not mention "Sage" — the
    underlying rubric JSONs sometimes include phrasings like
    "Sage-automated evaluation" in q2 notes or other free-text fields."""
    if not text:
        return text
    for old, new in _SAGE_REPLACEMENTS:
        text = text.replace(old, new)
    return text


def _load_mapping_table():
    """Load the question-mapping table that says which rubric questions
    also appear on the web form. Used to label each question with its
    origin (Rubric only / Web form only / Rubric + Web form)."""
    mpath = Path(__file__).resolve().parent / "mapping_table.json"
    if not mpath.exists():
        return {"science": [], "practitioner": []}
    return json.loads(mpath.read_text())


def _origin_label(spec, mapping_records):
    """Return a short (rubric / web form / both) tag for one question."""
    # Web-form-only entries carry an explicit source marker
    if spec.get("source") in ("science_form", "practitioner_form"):
        return "Web form"
    qid = spec["id"]
    # Look for a mapping record that lists this rubric id
    for m in mapping_records:
        if qid in (m.get("excel_rubric_ids") or []):
            if m.get("google_form_names"):
                return "Rubric + Web form"
            return "Rubric"
    return "Rubric"


def _rubric_section_flowables(styles, title, schema, rubric_data, anchor_prefix,
                               mapping_records=None):
    """Build the flowables for one rubric section. anchor_prefix is "s" or
    "p" so intra-PDF hyperlinks (e.g., 'Sources: s.q4') can target each
    question's anchor.

    Each question is rendered with its origin label, a Self-reported answer
    row (placeholder until web-form data arrives), and the materials-derived
    answer row.
    """
    out = []
    out.append(PageBreak())
    out.append(Paragraph(_escape(title), styles["h1"]))
    out.append(Spacer(1, 4))

    # Subhead with overview counts
    ans = rubric_data.get("answers", {})
    br = rubric_data.get("branching_notes", {})
    gaps = rubric_data.get("evidence_gaps", {})
    filled = sum(1 for v in ans.values() if v is not None)
    out.append(Paragraph(
        f"<b>{filled} answered</b> · {len(br)} branching (n/a) · "
        f"{len(gaps)} not addressed in available materials · "
        f"{len(schema)} total questions",
        styles["body_muted"],
    ))
    out.append(Paragraph(
        "For each question below, two voices are presented: the team's "
        "self-reported answer (from their web-form response) and the "
        "review-agent's extraction from the submission materials. Where "
        "a question is not asked on the web form, only the materials "
        "extraction is shown.",
        styles["body_muted"],
    ))
    out.append(Spacer(1, 10))

    mapping_records = mapping_records or []

    for spec in schema:
        qid = spec["id"]
        q_text = spec["text"]
        answer = ans.get(qid)
        origin = _origin_label(spec, mapping_records)

        # Anchor for intra-PDF hyperlinks. Format: "s_q4" or "p_q9bii".
        anchor_name = f"{anchor_prefix}_{qid}"
        # Append a small grey origin tag after the question text
        origin_tag = (f' <font size="8" color="#64748B">'
                      f'[{_escape(origin)}]</font>')
        q_para_text = (
            f'<a name="{anchor_name}"/>'
            f'<b>{qid.upper()}.</b> {_escape(q_text)}{origin_tag}'
        )
        block = [Paragraph(q_para_text, styles["question_text"])]

        # --- Voice 1: Team self-report (web form) ---
        # Only show this row if the question is asked on the web form.
        if origin in ("Web form", "Rubric + Web form"):
            block.append(Paragraph(
                "<b>Self-reported:</b> <i>Will be added when web-form data "
                "arrives (expected 2026-05-26).</i>",
                styles["notes"],
            ))

        # --- Voice 2: Review agent's extraction from materials ---
        # Always rendered. If the question is in the team_only bucket,
        # show the team_only note instead of attempting extraction.
        team_only = rubric_data.get("team_only", {}) or {}
        if qid in team_only:
            block.append(Paragraph(
                f"<b>Review agent:</b> <i>{_escape(_strip_sage_naming(_coerce_reason(team_only[qid])))}</i>",
                styles["notes"],
            ))
        elif answer is not None:
            block.append(Paragraph(
                "<b>Review agent:</b>", styles["answer_marker"],
            ))
            for style_key, text in _format_answer(spec, answer):
                block.append(Paragraph(text, styles[style_key]))
        elif qid in br:
            for style_key, text in _format_skipped(qid, br[qid], "branching"):
                block.append(Paragraph(
                    f"<b>Review agent:</b> {text}", styles[style_key],
                ))
        elif qid in gaps:
            for style_key, text in _format_skipped(qid, gaps[qid], "gap"):
                block.append(Paragraph(
                    f"<b>Review agent:</b> {text}", styles[style_key],
                ))
        else:
            # Defensive: shouldn't happen with the coverage check.
            block.append(Paragraph(
                "<b>Review agent:</b> <i>(no answer or explanation "
                "recorded — please re-run the rubric skill)</i>",
                styles["notes"],
            ))

        out.append(KeepTogether(block))

    return out


# ---------------------------------------------------------------------------
# Cover / Exec summary
# ---------------------------------------------------------------------------

def _human_date(iso_string):
    try:
        dt = datetime.fromisoformat(iso_string.replace("Z", "+00:00"))
        return dt.strftime("%B %-d, %Y · %H:%M UTC")
    except Exception:
        return iso_string


def _model_name(s_data, p_data):
    """Pull model name from q1 of either rubric — they should agree."""
    for d in (s_data, p_data):
        q1 = (d.get("answers", {}) or {}).get("q1") or {}
        if q1.get("value"):
            return q1["value"]
    return "(model name not set)"


def _location_from_ws(workspace_name):
    if "(Forest)" in workspace_name:
        return "Town of Forests"
    if "(Prairie)" in workspace_name:
        return "Town of Prairies"
    return "Location not detected from workspace name"


def _cover_flowables(styles, workspace_name, s_data, p_data):
    out = []
    out.append(Spacer(1, 1.2 * inch))
    out.append(Paragraph("Comprehensive Review", styles["cover_title"]))
    out.append(Spacer(1, 0.15 * inch))

    out.append(Paragraph(_escape(workspace_name), styles["cover_subtitle"]))
    out.append(Spacer(1, 0.4 * inch))

    out.append(Paragraph(f"<b>Model:</b> {_escape(_model_name(s_data, p_data))}", styles["cover_meta"]))
    out.append(Paragraph(f"<b>Location:</b> {_escape(_location_from_ws(workspace_name))}", styles["cover_meta"]))
    out.append(Paragraph(
        f"<b>Project:</b> Fire Risk Modeling Exercise — Deep Synthesis",
        styles["cover_meta"],
    ))
    out.append(Paragraph(
        f"<b>Generated:</b> {_human_date(datetime.now(timezone.utc).isoformat())}",
        styles["cover_meta"],
    ))

    out.append(Spacer(1, 1.0 * inch))

    # Headline finding box — counts at a glance
    s_filled = sum(1 for v in s_data.get("answers", {}).values() if v is not None)
    p_filled = sum(1 for v in p_data.get("answers", {}).values() if v is not None)
    s_gaps   = len(s_data.get("evidence_gaps", {}))
    p_gaps   = len(p_data.get("evidence_gaps", {}))

    headline_data = [[
        Paragraph(f"<b>Science Rubric</b><br/>{s_filled}/{len(SCIENCE_RUBRIC)} answered<br/>"
                  f"<font size=8 color='#9F1239'>{s_gaps} not addressed</font>",
                  styles["cover_meta"]),
        Paragraph(f"<b>Practitioner Rubric</b><br/>{p_filled}/{len(PRACTITIONER_RUBRIC)} answered<br/>"
                  f"<font size=8 color='#9F1239'>{p_gaps} not addressed</font>",
                  styles["cover_meta"]),
    ]]
    headline_table = Table(headline_data, colWidths=[3 * inch, 3 * inch])
    headline_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BACKGROUND", (0, 0), (-1, -1), COLOR_LIGHT_BG),
        ("BOX", (0, 0), (-1, -1), 0.5, COLOR_BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, COLOR_BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
    ]))
    out.append(headline_table)

    return out


def _exec_summary_table(s_data, p_data):
    """Build the Model At-a-Glance table data."""
    def val(d, qid, default="—"):
        a = (d.get("answers", {}) or {}).get(qid)
        if not a:
            return default
        v = a.get("value")
        sel = a.get("selected")
        if sel:
            return ", ".join(sel)
        return v or default

    def loc_coverage(ws_name):
        return "Both" if "(Both)" in ws_name else ("Forest only" if "(Forest)" in ws_name else
                                                    "Prairie only" if "(Prairie)" in ws_name else
                                                    "Unknown")

    rows = [
        ["Model name",         _model_name(s_data, p_data)],
        ["Type (static/dynamic)", val(s_data, "q4")],
        ["Approach",           val(s_data, "q18")],
        ["Output scale",       val(s_data, "q8")],
        ["Risk score type",    val(s_data, "q13a", default=val(s_data, "q13"))],
        ["Risk decomposition", val(s_data, "q17b", default=val(s_data, "q17"))],
        ["Includes ember model?",  val(s_data, "q4f")],
        ["Structure-to-structure?", val(s_data, "q9civ")],
        ["Primary use",        val(s_data, "q20", default=val(p_data, "q3"))],
        ["Secondary use",      val(s_data, "q20a", default=val(p_data, "q4"))],
        ["Decision-support clarity (practitioner)", val(p_data, "q4a")],
        ["Parcel-level insights (practitioner)",    val(p_data, "q7")],
    ]
    return rows


def _exec_summary_flowables(styles, s_data, p_data):
    out = []
    out.append(PageBreak())
    out.append(Paragraph("Executive Summary", styles["h1"]))

    # Render each source rubric note as its own subsection with a heading.
    # Concatenating raw notes (the prior approach) produced collisions when
    # each note carries its own numbered list — readers saw
    # "1. ... 2. ... 3. ... 1. ... 2. ... 3. ..." as one wall of text.
    summary_sources = [
        ("s", "q19", "What this model measures"),
        ("s", "q21", "Greatest strengths (science view)"),
        ("s", "q22", "Ensemble opportunity (science view)"),
        ("p", "q10",  "Top strengths (practitioner view)"),
        ("p", "q10a", "Greatest strengths (practitioner view)"),
        ("p", "q10b", "Ensemble opportunity (practitioner view)"),
    ]
    any_rendered = False
    for src, qid, heading in summary_sources:
        data = s_data if src == "s" else p_data
        a = (data.get("answers", {}) or {}).get(qid)
        if not a:
            continue
        v = a.get("value")
        notes = (a.get("notes") or "").strip()
        content = (notes or v or "").strip()
        if not content:
            continue
        out.append(Paragraph(_escape(heading), styles["h3"]))
        out.append(Paragraph(_escape(content), styles["body"]))
        any_rendered = True

    if not any_rendered:
        out.append(Paragraph(
            "No top-level synthesis recorded in the rubric notes. See the "
            "rubric and synthesis sections below for question-by-question detail.",
            styles["body"]))

    out.append(Spacer(1, 14))

    # Model at-a-glance table
    out.append(Paragraph("Model at a Glance", styles["h2"]))
    rows = _exec_summary_table(s_data, p_data)
    rendered = [[Paragraph(f"<b>{_escape(r[0])}</b>", getSampleStyleSheet()["BodyText"]),
                 Paragraph(_escape(r[1]), getSampleStyleSheet()["BodyText"])]
                for r in rows]
    table = Table(rendered, colWidths=[2.4 * inch, 4 * inch])
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND", (0, 0), (0, -1), COLOR_LIGHT_BG),
        ("BOX", (0, 0), (-1, -1), 0.5, COLOR_BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, COLOR_BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    out.append(table)
    return out


# ---------------------------------------------------------------------------
# Synthesis questions section
# ---------------------------------------------------------------------------

def _resolve_rubric_field(s_data, p_data, dotted_id):
    """Given a 's.q4' or 'p.q9bii' style id, return (rubric_label, spec, answer)."""
    prefix, qid = dotted_id.split(".", 1)
    if prefix == "s":
        data = s_data
        by_id = {q["id"]: q for q in SCIENCE_RUBRIC}
        label = "Science"
        anchor = f"s_{qid}"
    elif prefix == "p":
        data = p_data
        by_id = {q["id"]: q for q in PRACTITIONER_RUBRIC}
        label = "Practitioner"
        anchor = f"p_{qid}"
    else:
        return None
    spec = by_id.get(qid)
    answer = (data.get("answers", {}) or {}).get(qid)
    return (label, spec, answer, anchor)


def _compose_synthesized(s_data, p_data, sources):
    """Compose a synthesized answer paragraph from multiple rubric fields.
    Returns (paragraph_text, list_of_anchor_links_for_attribution)."""
    parts = []
    used_anchors = []
    for src in sources:
        resolved = _resolve_rubric_field(s_data, p_data, src)
        if resolved is None:
            continue
        label, spec, answer, anchor = resolved
        if spec is None:
            continue
        if answer is None:
            continue  # skip nulls
        # Build a short factual line
        val = answer.get("value")
        sel = answer.get("selected")
        notes = (answer.get("notes") or "").strip()
        bullet = ""
        if val:
            bullet = f"<b>{src.upper()}:</b> {_escape(val)}"
        elif sel:
            bullet = f"<b>{src.upper()}:</b> {_escape(', '.join(sel))}"
        if notes:
            bullet = (bullet + " — " if bullet else f"<b>{src.upper()}:</b> ") + _escape(notes)
        if bullet:
            parts.append(f'<a href="#{anchor}">{bullet}</a>')
            used_anchors.append(src)
    if not parts:
        return None, []
    composed = "<br/><br/>".join(parts)
    return composed, used_anchors


def _direct_answer(s_data, p_data, sources):
    """For 'direct' answer_kind: pull the first source that has a value
    and return its representation."""
    for src in sources:
        resolved = _resolve_rubric_field(s_data, p_data, src)
        if resolved is None:
            continue
        label, spec, answer, anchor = resolved
        if not answer:
            continue
        v = answer.get("value")
        sel = answer.get("selected")
        notes = (answer.get("notes") or "").strip()
        if v or sel or notes:
            return _compose_synthesized(s_data, p_data, sources)
    return None, []


def _synthesis_flowables(styles, s_data, p_data, synthesis_data):
    out = []
    out.append(PageBreak())
    out.append(Paragraph("Synthesis Questions", styles["h1"]))
    out.append(Paragraph(
        "These questions are from the boss's reviewer worksheet (per-submission "
        "scope). Most answers are composed from the Science and/or Practitioner "
        "rubric. Click any blue rubric reference to jump to the source question.",
        styles["body_muted"],
    ))
    out.append(Spacer(1, 8))

    for section in synthesis_data["sections"]:
        # Filter to per-submission questions only
        per_sub = [q for q in section["questions"] if q["scope"] == "per_submission"]
        if not per_sub:
            continue
        out.append(Paragraph(_escape(section["name"]), styles["h2"]))
        for q in per_sub:
            block = [Paragraph(_escape(q["text"]), styles["question_text"])]
            if q["answer_kind"] == "no_rubric_basis":
                block.append(Paragraph(
                    "Requires reviewer judgment — this question is not directly "
                    "answered by the Science or Practitioner rubric.",
                    styles["gap_marker"],
                ))
            else:
                sources = q.get("answers_from", [])
                composed, used = _compose_synthesized(s_data, p_data, sources)
                if composed:
                    block.append(Paragraph(composed, styles["answer_text"]))
                    if used:
                        attr = "Sources: " + ", ".join(
                            f'<a href="#{src.replace(".", "_")}">{src.upper()}</a>'
                            for src in used
                        )
                        block.append(Paragraph(attr, styles["source_attribution"]))
                else:
                    block.append(Paragraph(
                        "No rubric answers available for the source fields. "
                        "Requires reviewer judgment.",
                        styles["gap_marker"],
                    ))
            out.append(KeepTogether(block))

    return out


# ---------------------------------------------------------------------------
# Placeholder sections for data sources not yet available
# ---------------------------------------------------------------------------

def _placeholder_section(styles, title, body_text):
    """Render a section heading + a muted-italic placeholder paragraph.
    Used for data sources that aren't yet integrated (team web-form
    responses, reviewer commentary, ArcGIS summary). Once the data
    arrives, replace the placeholder body with actual content."""
    out = []
    out.append(PageBreak())
    out.append(Paragraph(_escape(title), styles["h1"]))
    out.append(Spacer(1, 4))
    out.append(Paragraph(
        f"<i>{_escape(body_text)}</i>",
        styles["body_muted"],
    ))
    return out


def _team_self_report_flowables(styles, s_data, p_data):
    return _placeholder_section(
        styles,
        "Section 4: Team Self-Report (Web Form)",
        "Each submitting team fills out two Survey123 forms (Science Evaluator "
        "and Practitioner Evaluator) describing their model in their own words. "
        "When those responses are delivered, they will be presented here next to "
        "the review agent's independent extraction from the submission materials. Until then, "
        "this section is a placeholder — only the materials-derived voice is "
        "available."
    )


def _reviewer_commentary_flowables(styles):
    return _placeholder_section(
        styles,
        "Section 5: Reviewer Commentary",
        "Reviewer commentary on this submission will be added here when received "
        "from the project manager. Until then, this section is empty."
    )


def _arcgis_summary_flowables(styles, ws_root):
    """If the workspace has an _arcgis_endpoints.txt file, list its contents
    here so reviewers see what services were exposed. Otherwise show a
    placeholder."""
    out = []
    out.append(PageBreak())
    out.append(Paragraph("Appendix: ArcGIS Service Summary", styles["h1"]))
    out.append(Spacer(1, 4))
    arcgis_file = ws_root / "_arcgis_endpoints.txt"
    if arcgis_file.exists():
        try:
            text = arcgis_file.read_text().strip()
        except Exception:
            text = ""
        if text:
            out.append(Paragraph(
                "The team's submission referenced the following ArcGIS endpoints. "
                "An external analyst summary of the layer contents will be "
                "added here when available.",
                styles["body_muted"],
            ))
            out.append(Spacer(1, 8))
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                # Format: "<label>\t<url>"
                if "\t" in line:
                    label, url = line.split("\t", 1)
                    out.append(Paragraph(
                        f"<b>{_escape(label)}</b> — <font face='Helvetica' size='9'>{_escape(url)}</font>",
                        styles["body"],
                    ))
                else:
                    out.append(Paragraph(_escape(line), styles["body"]))
            return out
    out.append(Paragraph(
        "<i>No ArcGIS endpoints were associated with this submission, OR an "
        "external analyst summary has not yet been provided. When an analyst "
        "summarizes the published ArcGIS layers, the summary will appear here.</i>",
        styles["body_muted"],
    ))
    return out


# ---------------------------------------------------------------------------
# Appendix — file inventory
# ---------------------------------------------------------------------------

def _appendix_flowables(styles, ws_root):
    out = []
    out.append(PageBreak())
    out.append(Paragraph("Appendix: File Inventory", styles["h1"]))

    mf_path = ws_root / "_manifest.json"
    if not mf_path.exists():
        out.append(Paragraph("No _manifest.json found in workspace.", styles["body_muted"]))
        return out

    mf = json.loads(mf_path.read_text())
    dl = mf.get("downloaded", [])
    sk = mf.get("skipped", [])
    er = mf.get("errors", [])

    out.append(Paragraph(
        f"Downloaded: <b>{len(dl)}</b> files · Skipped: <b>{len(sk)}</b> · Errors: <b>{len(er)}</b>",
        styles["body"],
    ))
    out.append(Spacer(1, 8))

    # By-category counts
    from collections import Counter
    by_cat = Counter(d.get("category", "?") for d in dl)
    if by_cat:
        out.append(Paragraph("Downloaded by category:", styles["h3"]))
        for k, v in by_cat.most_common():
            out.append(Paragraph(f"  • {_escape(k)}: {v}", styles["body"]))

    by_reason = Counter(s.get("reason", "?") for s in sk)
    if by_reason:
        out.append(Spacer(1, 8))
        out.append(Paragraph("Skipped by reason:", styles["h3"]))
        for k, v in by_reason.most_common():
            out.append(Paragraph(f"  • {_escape(k)}: {v}", styles["body"]))

    if er:
        out.append(Spacer(1, 8))
        out.append(Paragraph("Errors:", styles["h3"]))
        for e in er[:10]:
            msg = e.get("error", "(no message)")[:200]
            out.append(Paragraph(f"  • {_escape(msg)}", styles["body_muted"]))
        if len(er) > 10:
            out.append(Paragraph(f"  … and {len(er) - 10} more", styles["body_muted"]))

    return out


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

class _NumberedCanvas(rl_canvas.Canvas):
    """Two-pass canvas that knows the total page count by deferring
    page commit until save time. Standard reportlab pattern for the
    "Page X of Y" idiom — saveState() each completed page, then on
    save() draw all pages knowing the total. Workspace name is captured
    when the canvas is created so we can put it in the footer."""

    def __init__(self, *args, workspace_name="", **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        self._workspace_name = workspace_name

    def showPage(self):
        # Capture the current page state instead of emitting it.
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        # Now we know the total. Emit each captured page with a proper
        # "Page X of Y" footer on every page except the cover (page 1).
        total = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            page_num = self._pageNumber
            if page_num > 1:
                # Repaint the right-side footer text with the correct total.
                self.saveState()
                self.setFont(FONT_BODY, 8)
                self.setFillColor(COLOR_MUTED)
                self.drawRightString(
                    PAGE_W - MARGIN_X, 0.4 * inch,
                    f"Page {page_num} of {total}"
                )
                self.restoreState()
            super().showPage()
        super().save()


class _DocTemplate(BaseDocTemplate):
    """Plain BaseDocTemplate — page-count handling is done by NumberedCanvas."""
    def __init__(self, filename, workspace_name, **kw):
        super().__init__(filename, **kw)
        self._workspace_name = workspace_name


def generate_report(ws_root, output_path=None):
    """Build the comprehensive PDF report for one workspace.

    Required: the workspace folder must contain BOTH
    `_rubric_science.json` and `_rubric_practitioner.json`. Missing
    either raises a clear ValueError with instructions.

    Returns the path of the written PDF.
    """
    ws_root = Path(ws_root)
    if not ws_root.is_dir():
        raise ValueError(f"Workspace folder not found: {ws_root}")

    science_path = ws_root / "_rubric_science.json"
    practitioner_path = ws_root / "_rubric_practitioner.json"

    missing = []
    if not science_path.exists():
        missing.append(("_rubric_science.json",
                        "Run: %%ask Generate a Science Evaluator rubric review for "
                        f"the \"{ws_root.name}\" workspace."))
    if not practitioner_path.exists():
        missing.append(("_rubric_practitioner.json",
                        "Run: %%ask Generate a Practitioner Evaluator rubric review for "
                        f"the \"{ws_root.name}\" workspace."))
    if missing:
        msg = ["Cannot generate comprehensive report. Required inputs missing:"]
        for fname, action in missing:
            msg.append(f"  ✗ {ws_root / fname}")
            msg.append(f"    -> {action}")
        msg.append("\nRun the missing rubric(s) first, then re-run this cell.")
        raise ValueError("\n".join(msg))

    s_data = json.loads(science_path.read_text())
    p_data = json.loads(practitioner_path.read_text())

    syn_path = Path(__file__).resolve().parent / "synthesis_questions.json"
    if not syn_path.exists():
        raise FileNotFoundError(
            f"synthesis_questions.json missing at {syn_path} — "
            "skill installation is incomplete."
        )
    synthesis_data = json.loads(syn_path.read_text())

    # Output path
    if output_path is None:
        ws_safe = re.sub(r'[\\/*?:"<>|]', "_", ws_root.name).strip("._-") or "submission"
        output_path = ws_root / f"_comprehensive_report_{ws_safe}.pdf"
    output_path = Path(output_path)

    styles = _styles()

    # Build the document; NumberedCanvas handles two-pass "Page X of Y"
    # via the standard reportlab pattern (capture page states, emit at save).
    doc = _DocTemplate(
        str(output_path),
        workspace_name=ws_root.name,
        pagesize=letter,
        leftMargin=MARGIN_X, rightMargin=MARGIN_X,
        topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM,
        title=f"Comprehensive Review — {ws_root.name}",
    )

    cover_frame = Frame(MARGIN_X, MARGIN_BOTTOM,
                        PAGE_W - 2 * MARGIN_X, PAGE_H - MARGIN_TOP - MARGIN_BOTTOM,
                        id="cover", showBoundary=0)
    content_frame = Frame(MARGIN_X, MARGIN_BOTTOM,
                          PAGE_W - 2 * MARGIN_X, PAGE_H - MARGIN_TOP - MARGIN_BOTTOM,
                          id="content", showBoundary=0)

    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[cover_frame],
                     onPage=_draw_cover_decoration),
        PageTemplate(id="content", frames=[content_frame],
                     onPage=lambda c, d: _draw_header_footer(c, d, ws_root.name)),
    ])

    story = []
    story.extend(_cover_flowables(styles, ws_root.name, s_data, p_data))

    # Switch templates for everything after the cover
    from reportlab.platypus import NextPageTemplate
    story.append(NextPageTemplate("content"))

    mapping = _load_mapping_table()
    story.extend(_exec_summary_flowables(styles, s_data, p_data))
    story.extend(_rubric_section_flowables(
        styles, "Section 1: Science Evaluator Rubric",
        SCIENCE_RUBRIC, s_data, "s",
        mapping_records=mapping.get("science", []),
    ))
    story.extend(_rubric_section_flowables(
        styles, "Section 2: Practitioner Evaluator Rubric",
        PRACTITIONER_RUBRIC, p_data, "p",
        mapping_records=mapping.get("practitioner", []),
    ))
    story.extend(_synthesis_flowables(styles, s_data, p_data, synthesis_data))
    story.extend(_reviewer_commentary_flowables(styles))
    story.extend(_arcgis_summary_flowables(styles, ws_root))
    story.extend(_appendix_flowables(styles, ws_root))

    # Build with NumberedCanvas — captures page states, then on save()
    # writes each page with the correct "Page X of Y" total.
    def _make_canvas(*args, **kwargs):
        return _NumberedCanvas(*args, workspace_name=ws_root.name, **kwargs)

    doc.build(story, canvasmaker=_make_canvas)
    return output_path
