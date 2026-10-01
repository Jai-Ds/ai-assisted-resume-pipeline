# -*- coding: utf-8 -*-
"""
build_resume.py — Universal flexible resume builder

Accepts a content JSON dict and renders a 1-page DOCX + PDF with auto-scale.
Name, contact, and education come from config.json (see config.example.json).

FIXED (never changes):    Name, contact, LinkedIn/GitHub, education (all from
                          config.json).
VARIABLE (from JSON):     Summary, bullets (any count), experience roles
                          (titles/dates from config.json, optional roles
                          included only when relevant — see job_runner.py),
                          projects (drop/swap/invent), competency categories
                          and labels.

Usage from job_runner.py:
    from build_resume import auto_build
    docx_path, pdf_path, fs = auto_build(content_dict, out_dir, out_name)

CLI usage:
    python build_resume.py content.json out_dir out_name

Expected JSON structure:
{
  "summary": "...",
  "experience_header": "PROFESSIONAL EXPERIENCE (AROUND 5 YEARS)",  # optional
  "experience": [
    {
      "title": "Product Manager-II, o9 Solutions",
      "dates": "May 2023 - Present (34 months)",
      "bullets": ["...", "...", "..."]       # any count
    },
    {
      "title": "Product Intern, Tenzai (AI Startup)",  # optional — omit to drop
      "dates": "Mar 2022 - Jun 2022 (3 Months)",
      "bullets": ["...", "..."]
    },
    {
      "title": "Program Analyst, Cognizant",
      "dates": "Dec 2017 - Sep 2019 (21 Months)",
      "bullets": ["...", "..."]
    }
  ],
  "projects_section_label": "KEY PROJECTS AND CAPSTONE WORK",  # optional
  "projects": [                                                  # optional — omit to skip section
    {"lead": "AI & ML Capstone Projects", "text": "...", "has_github_link": true},
    {"lead": "Enterprise Planning Templates (o9)", "text": "..."}
  ],
  "competencies": [                                              # any count/label
    {"label": "Product & Supply Chain", "items": "..."},
    {"label": "Strategy & Research",    "items": "..."},
    {"label": "Tools & Tech",           "items": "..."}
  ]
}

FORMATTING RULES (critical — do not change):
  - Font        : Nunito
  - Margins     : Top 0.55in, Bottom 0.50in, Left/Right 0.65in
  - Section hdr : Bold, gray #C0C0C0 background
  - Bold inline : **text** markers via rr() — never en dash inside bold runs
  - Auto-scale  : binary search 9.5pt → 13.0pt to fill exactly 1 page
"""

import copy
import os
import re
import sys

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx2pdf import convert
import pdfplumber

from config import PROFILE

HEADER_SHADE = "C0C0C0"
FONT         = "Nunito"
RIGHT_TAB    = 10080
NAME         = PROFILE["name"]
LOCATION     = PROFILE["location"]
EMAIL        = PROFILE["email"]
PHONE        = PROFILE["phone"]
LINKEDIN_URL = PROFILE["linkedin_url"]
GITHUB_URL   = PROFILE["github_url"]


# ── Spacing (all proportional to font size) ────────────────────────────────────
def SP(fs):
    k = fs / 10.0
    return dict(
        name_after    = 1.0 * k,
        contact_after = 3.0 * k,
        sec_before    = 3.5 * k,
        sec_after     = 1.5 * k,
        job_before    = 3.0 * k,
        job_after     = 0.5 * k,
        bullet_line   = fs * 1.13,
        bullet_after  = 0.8 * k,
        edu_before    = 2.0 * k,
        edu_after     = 1.0 * k,
        sum_after     = 2.0 * k,
        sum_line      = fs * 1.15,
    )


# ── XML helpers ────────────────────────────────────────────────────────────────
def shade(p):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), HEADER_SHADE)
    pPr.append(shd)


def rtab(p):
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tab  = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), str(RIGHT_TAB))
    tabs.append(tab)
    pPr.append(tabs)


def sp(p, before=0, after=0, line=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)
    if line:
        pf.line_spacing = Pt(line)


def R(p, text, fs, bold=False):
    r = p.add_run(text)
    r.font.name      = FONT
    r.font.size      = Pt(fs)
    r.font.bold      = bold
    r.font.color.rgb = RGBColor(0, 0, 0)
    r.font.underline = False
    return r


def RL(p, text, url, fs):
    """Render a clickable hyperlink run."""
    part = p.part
    rid  = part.relate_to(
        url,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        is_external=True,
    )
    hl   = OxmlElement('w:hyperlink')
    hl.set(qn('r:id'), rid)
    nr   = OxmlElement('w:r')
    rPr  = OxmlElement('w:rPr')
    rf   = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), FONT)
    rf.set(qn('w:hAnsi'), FONT)
    rPr.append(rf)
    sz  = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(fs * 2)))
    rPr.append(sz)
    col = OxmlElement('w:color')
    col.set(qn('w:val'), '0563C1')
    rPr.append(col)
    u   = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    nr.append(rPr)
    t      = OxmlElement('w:t')
    t.text = text
    nr.append(t)
    hl.append(nr)
    p._p.append(hl)


def rr(p, text, fs):
    """Render text with **bold** markers. Never use en dash inside bold runs."""
    for i, part in enumerate(re.split(r'\*\*(.*?)\*\*', text)):
        if part:
            R(p, part, fs, bold=(i % 2 == 1))


def margins(doc):
    s = doc.sections[0]
    s.top_margin    = Inches(0.55)
    s.bottom_margin = Inches(0.50)
    s.left_margin   = Inches(0.65)
    s.right_margin  = Inches(0.65)


# ── Block renderers ────────────────────────────────────────────────────────────
def name_line(doc, fs, s):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp(p, after=s['name_after'])
    R(p, NAME.upper(), fs + 8, bold=True)


def contact(doc, fs, s):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sp(p, after=s['contact_after'])
    R(p, f"{LOCATION}  |  {EMAIL}  |  {PHONE}  |  ", fs)
    RL(p, "LinkedIn Profile", LINKEDIN_URL, fs)


def hdr(doc, title, fs, s):
    p = doc.add_paragraph()
    sp(p, before=s['sec_before'], after=s['sec_after'])
    shade(p)
    p.paragraph_format.left_indent = Inches(0.05)
    R(p, f"  {title}", fs, bold=True)


def job_line(doc, title_co, dates, fs, s):
    p = doc.add_paragraph()
    sp(p, before=s['job_before'], after=s['job_after'])
    rtab(p)
    R(p, title_co, fs, bold=True)
    tr = p.add_run("\t")
    tr.font.name = FONT
    tr.font.size = Pt(fs)
    R(p, dates, fs, bold=True)


def bul(doc, text, fs, s, lead=None):
    """Bullet point. lead (if given) renders bold before the text."""
    p = doc.add_paragraph()
    sp(p, after=s['bullet_after'], line=s['bullet_line'])
    p.paragraph_format.left_indent       = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.18)
    R(p, "\u2022  ", fs)
    if lead:
        R(p, lead, fs, bold=True)
    rr(p, text, fs)


def bul_link(doc, lead, text, link_text, url, fs, s):
    """Bullet with bold lead, rr() body text, and a trailing hyperlink."""
    p = doc.add_paragraph()
    sp(p, after=s['bullet_after'], line=s['bullet_line'])
    p.paragraph_format.left_indent       = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.18)
    R(p, "\u2022  ", fs)
    if lead:
        R(p, lead, fs, bold=True)
    rr(p, text, fs)
    RL(p, link_text, url, fs)


def edu_e(doc, deg, dates, school, fs, s):
    p = doc.add_paragraph()
    sp(p, before=s['edu_before'])
    rtab(p)
    p.paragraph_format.left_indent = Inches(0.1)
    R(p, deg, fs, bold=True)
    tr = p.add_run("\t")
    tr.font.name = FONT
    tr.font.size = Pt(fs)
    R(p, dates, fs, bold=True)
    p2 = doc.add_paragraph()
    sp(p2, after=s['edu_after'])
    p2.paragraph_format.left_indent = Inches(0.1)
    R(p2, school, fs)


def comp_line(doc, cat, items, fs, s):
    p = doc.add_paragraph()
    sp(p, after=s['bullet_after'], line=s['bullet_line'])
    p.paragraph_format.left_indent       = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.18)
    R(p, "\u2022  ", fs)
    R(p, f"{cat}: ", fs, bold=True)
    R(p, items, fs)


def summ(doc, text, fs, s):
    p = doc.add_paragraph()
    sp(p, after=s['sum_after'], line=s['sum_line'])
    rr(p, text, fs)


# ── Core build ─────────────────────────────────────────────────────────────────
def build(fs, out_docx, content):
    """Render the resume at font size fs and save to out_docx."""
    doc = Document()
    margins(doc)
    s = SP(fs)

    # ── Fixed header ──────────────────────────────────────────────────────────
    name_line(doc, fs, s)
    contact(doc, fs, s)

    # ── Summary (fully rewritten by claude) ───────────────────────────────────
    hdr(doc, "SUMMARY", fs, s)
    summ(doc, content["summary"], fs, s)

    # ── Experience (variable roles + bullet counts) ───────────────────────────
    exp_header = content.get("experience_header", "PROFESSIONAL EXPERIENCE (AROUND 5 YEARS)")
    hdr(doc, exp_header, fs, s)
    for exp in content.get("experience", []):
        job_line(doc, exp["title"], exp["dates"], fs, s)
        for bullet in exp.get("bullets", []):
            bul(doc, bullet, fs, s)

    # ── Projects (optional — drop entire section if empty) ────────────────────
    projects = content.get("projects") or []
    if projects:
        proj_label = content.get("projects_section_label", "KEY PROJECTS AND CAPSTONE WORK")
        hdr(doc, proj_label, fs, s)
        for proj in projects:
            lead_raw  = proj.get("lead", "")
            body_text = proj.get("text", "")
            # Normalise lead: ensure ": " suffix
            lead_str  = None
            if lead_raw:
                lead_str = lead_raw.rstrip(": ") + ": "
            if proj.get("has_github_link"):
                body_with_space = body_text if body_text.endswith(" ") else body_text + " "
                bul_link(doc, lead_str or "", body_with_space, "GitHub- Repo Link", GITHUB_URL, fs, s)
            else:
                bul(doc, body_text, fs, s, lead=lead_str)

    # ── Education (fixed — never changes; from config.json) ──────────────────
    hdr(doc, "EDUCATION", fs, s)
    for edu in PROFILE["education"]:
        edu_e(doc, edu["degree"], edu["dates"], edu["school"], fs, s)

    # ── Competencies (variable count + labels — reordered per JD) ────────────
    hdr(doc, "CORE COMPETENCIES", fs, s)
    for c in content.get("competencies", []):
        comp_line(doc, c["label"], c["items"], fs, s)

    doc.save(out_docx)


# ── Auto-scale ─────────────────────────────────────────────────────────────────
def count_pages(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return len(pdf.pages)


# Floors so trimming never guts the sections that carry the most signal.
MIN_O9_BULLETS   = 3
MIN_COMPETENCIES = 3


def _trim_steps(content):
    """
    Build an ordered list of trim actions, weakest content first.

    job_runner's tailoring prompt already orders each list (bullets,
    projects, competencies) most-to-least relevant to the JD and puts the
    least essential experience block (Cognizant) last, so "trim from the
    end" doubles as "cut what matters least for this JD/role" — no extra
    schema needed. Each action mutates `content` in place and returns a
    description of what it cut, or None if there was nothing left to cut.
    """
    def drop_last_project():
        projects = content.get("projects") or []
        if projects:
            removed = projects.pop()
            return f"dropped project '{removed.get('lead', '').strip(': ')}'"
        return None

    def drop_last_bullet_from(role_keyword, floor=0):
        for exp in content.get("experience", []):
            if role_keyword in exp.get("title", ""):
                bullets = exp.get("bullets", [])
                if len(bullets) > floor:
                    removed = bullets.pop()
                    return f"dropped {role_keyword} bullet: \"{removed[:60]}...\""
        return None

    def drop_empty_role(role_keyword):
        exps = content.get("experience", [])
        for i, exp in enumerate(exps):
            if role_keyword in exp.get("title", "") and not exp.get("bullets"):
                exps.pop(i)
                return f"dropped now-empty {role_keyword} block"
        return None

    def drop_last_competency():
        comps = content.get("competencies", [])
        if len(comps) > MIN_COMPETENCIES:
            removed = comps.pop()
            return f"dropped competency '{removed.get('label', '')}'"
        return None

    return [
        drop_last_project,
        drop_last_project,
        lambda: drop_last_bullet_from("Cognizant"),
        lambda: drop_last_bullet_from("Tenzai"),
        lambda: drop_last_bullet_from("Tenzai"),
        lambda: drop_empty_role("Tenzai"),
        drop_last_competency,
        lambda: drop_last_bullet_from("o9 Solutions", floor=MIN_O9_BULLETS),
        lambda: drop_last_bullet_from("o9 Solutions", floor=MIN_O9_BULLETS),
    ]


def auto_build(content, out_dir, out_name, min_fs=9.5, max_fs=13.0):
    """
    Mandatory 1-page enforcement: binary-search font size to fill exactly one
    PDF page; if content still overflows at the minimum font, progressively
    trim the lowest-priority content (see _trim_steps) and retry, escalating
    one cut at a time until it fits. Raises if trimming is exhausted and it
    still doesn't fit — a >1 page resume is never returned silently.

    Returns (docx_path, pdf_path, font_size).
    """
    os.makedirs(out_dir, exist_ok=True)
    docx_path = os.path.join(out_dir, out_name + ".docx")
    pdf_path  = os.path.join(out_dir, out_name + ".pdf")

    working = copy.deepcopy(content)
    steps   = _trim_steps(working)
    trims_applied = []

    for step in [None] + steps:
        if step is not None:
            desc = step()
            if desc is None:
                continue
            trims_applied.append(desc)
            print(f"  TRIM: {desc}")
        build(min_fs, docx_path, working)
        convert(docx_path, pdf_path)
        if count_pages(pdf_path) == 1:
            break
    else:
        raise RuntimeError(
            f"Resume still overflows past 1 page at {min_fs}pt after exhausting "
            f"all content trims ({len(trims_applied)} cuts applied): {trims_applied}. "
            f"Shorten the tailored content JSON by hand."
        )

    if trims_applied:
        print(f"  Applied {len(trims_applied)} content trim(s) to fit 1 page.")

    # Binary search: find largest font that still fits 1 page, on the
    # (possibly trimmed) content.
    lo, hi = min_fs, max_fs
    for _ in range(9):
        mid = round((lo + hi) / 2, 3)
        build(mid, docx_path, working)
        convert(docx_path, pdf_path)
        if count_pages(pdf_path) == 1:
            lo = mid
        else:
            hi = mid

    # Final render at best font size
    build(lo, docx_path, working)
    convert(docx_path, pdf_path)
    print(f"  Font  : {lo:.2f}pt  |  Pages: {count_pages(pdf_path)}")
    print(f"  DOCX  : {docx_path}")
    print(f"  PDF   : {pdf_path}")
    return docx_path, pdf_path, lo


# ── CLI entry ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import json
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 4:
        print("Usage: python build_resume.py content.json out_dir out_name")
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        _content = json.load(f)
    auto_build(_content, sys.argv[2], sys.argv[3])
