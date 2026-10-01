# -*- coding: utf-8 -*-
"""
job_runner.py — Resume Tailoring Runner (JSON architecture)

Reads Daily_Jobs_Log.xlsx, finds rows where:
  - Status = "Shortlisted"
  - Runner Status is blank

For each qualifying job:
  1. Reads all 4 base DOCX resumes at runtime from Base Versions/
  2. Reads Work DNA from memory file at runtime
  3. One claude -p call returns tailored content JSON + ATS analysis
  4. Calls build_resume.auto_build() → DOCX + PDF in Job Applications/
  5. Calls job_tracker.finalize() → ATS Excel + master tracker
  6. Updates Excel: Base Resume, Runner Status, Fit Score

Run manually after reviewing and shortlisting jobs in Daily_Jobs_Log.xlsx.
"""

import os
import sys
import re
import json
import subprocess
import tempfile
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from docx import Document

sys.stdout.reconfigure(encoding="utf-8")


# ── Config ────────────────────────────────────────────────────────────────────
# Your profile, paths, and base-resume tags all come from config.json (copy
# config.example.json to get started) — nothing below is hardcoded to one person.

from config import PROFILE, PATHS, BASE_RESUME_TAGS

PIPELINE_DIR   = PATHS["pipeline_dir"]
DAILY_LOG_FILE = os.path.join(PIPELINE_DIR, "Daily_Jobs_Log.xlsx")

JOB_APPS_DIR   = PATHS["job_apps_dir"]
BASE_VERS_DIR  = PATHS["base_versions_dir"]
WORK_DNA_FILE  = PATHS["work_dna_file"]

# Add Job Applications to path for job_tracker import
sys.path.insert(0, JOB_APPS_DIR)
# Add Job Pipeline to path for build_resume import
sys.path.insert(0, PIPELINE_DIR)

# Column indices in Daily_Jobs_Log.xlsx (1-based)
COL_COMPANY     = 2
COL_TITLE       = 3
COL_URL         = 5
COL_JD          = 6
COL_STATUS      = 7
COL_BASE_RESUME = 10
COL_RUNNER_ST   = 11
COL_FIT_SCORE   = 12

HEADER_ROW = 2   # row 1 = title banner, row 2 = column headers, row 3+ = data

MID_BLUE = "2E75B6"
WHITE    = "FFFFFF"

# Base resume files (tag → filename without extension) — from config.json
BASE_RESUME_FILES = BASE_RESUME_TAGS


# ── Helpers ───────────────────────────────────────────────────────────────────

def sanitize(text, maxlen=25):
    """Alphanumeric + underscores only, trimmed to maxlen."""
    return re.sub(r'[^\w]', '', (text or "").replace(' ', '_'))[:maxlen]


# ── Base resume reader ────────────────────────────────────────────────────────

def read_docx_text(path):
    """Extract all paragraph text from a DOCX file."""
    try:
        doc = Document(path)
        lines = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(lines)
    except Exception as e:
        print(f"  WARNING: Could not read {path}: {e}")
        return ""


def load_all_base_resumes():
    """
    Load all 4 base resumes from Base Versions/.
    Returns dict: tag -> text (or empty string if file missing).
    """
    resumes = {}
    for tag, fname in BASE_RESUME_FILES.items():
        fpath = os.path.join(BASE_VERS_DIR, fname + ".docx")
        if os.path.exists(fpath):
            resumes[tag] = read_docx_text(fpath)
        else:
            print(f"  WARNING: Base resume not found: {fpath}")
            resumes[tag] = ""
    return resumes


def load_work_dna():
    """Load Work DNA markdown from memory file."""
    try:
        with open(WORK_DNA_FILE, encoding="utf-8") as f:
            text = f.read()
        # Strip YAML frontmatter if present
        if text.startswith("---"):
            end = text.find("---", 3)
            if end != -1:
                text = text[end + 3:].strip()
        return text
    except Exception as e:
        print(f"  WARNING: Could not read Work DNA: {e}")
        return ""


# ── Excel helpers ─────────────────────────────────────────────────────────────

def ensure_runner_columns(ws):
    """
    Ensure Base Resume (col 10), Runner Status (col 11), Fit Score (col 12)
    header cells exist. Safe to call on any existing file.
    """
    col_defs = [
        (COL_BASE_RESUME, "Base Resume",   14),
        (COL_RUNNER_ST,   "Runner Status", 14),
    ]
    for col, label, width in col_defs:
        cell = ws.cell(HEADER_ROW, col)
        if cell.value != label:
            cell.value     = label
            cell.font      = Font(bold=True, color=WHITE, size=10)
            cell.fill      = PatternFill("solid", fgColor=MID_BLUE)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(col)].width = width

    # Runner Status dropdown
    dv = DataValidation(
        type="list",
        formula1='"Done,Failed,Skipped-NoJD"',
        allow_blank=True,
        showDropDown=False,
    )
    ws.add_data_validation(dv)
    for row in ws.iter_rows(min_row=HEADER_ROW + 1):
        dv.add(row[COL_RUNNER_ST - 1])


def find_shortlisted(ws):
    """Return list of (row_num, company, title, jd, url) for Shortlisted + blank Runner Status."""
    results = []
    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=False):
        status    = row[COL_STATUS - 1].value
        runner_st = row[COL_RUNNER_ST - 1].value
        if status == "Shortlisted" and not runner_st:
            company = row[COL_COMPANY - 1].value or ""
            title   = row[COL_TITLE   - 1].value or ""
            jd      = row[COL_JD      - 1].value or ""
            url     = row[COL_URL     - 1].value or ""
            results.append((row[0].row, company, title, jd, url))
    return results


def set_runner_status(ws, row_num, base_tag, status):
    """Write Base Resume and Runner Status into the Excel row."""
    for col, val in [(COL_BASE_RESUME, base_tag), (COL_RUNNER_ST, status)]:
        cell           = ws.cell(row_num, col)
        cell.value     = val
        cell.font      = Font(size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ── Claude call ───────────────────────────────────────────────────────────────
# The prompt is assembled dynamically from config.json's `profile` block, so it
# works for any number of roles/education entries and any role type (PM, sales,
# hardware, whatever) — nothing below names a specific person or role.

def _build_fixed_elements(profile):
    lines = [
        f'- Name: {profile["name"]}',
        f'- Contact: {profile["location"]} | {profile["email"]} | {profile["phone"]} | LinkedIn Profile',
    ]
    for role in profile["fixed_roles"]:
        tag = " (if included)" if role.get("optional") else ""
        lines.append(f'- Role title{tag}: "{role["title"]}"')
        lines.append(f'  - Dates: "{role["dates"]}"')
    for edu in profile["education"]:
        lines.append(f'- Education: {edu["degree"]}, {edu["dates"]}, {edu["school"]}')
    return "\n".join(lines)


def _build_bullet_limits(profile):
    lines = []
    for role in profile["fixed_roles"]:
        opt = " (if included)" if role.get("optional") else ""
        lines.append(f'   - {role["title"]}{opt}: max {role["max_bullets"]} bullets')
    lines.append("   - Projects section: max 2 items, keep each description under 2 lines")
    lines.append("   - Summary: max 3 lines when rendered")
    lines.append("   - Competencies: max 5 categories")
    return "\n".join(lines)


def _build_experience_schema_example(profile):
    entries = []
    for role in profile["fixed_roles"]:
        entries.append(
            '    {\n'
            f'      "title": "{role["title"]}",\n'
            f'      "dates": "{role["dates"]}",\n'
            '      "bullets": ["bullet 1", "bullet 2", "..."]\n'
            '    }'
        )
    return "[\n" + ",\n".join(entries) + "\n  ]"


def _build_hard_limits_line(profile):
    parts = [f'{role["title"]} max {role["max_bullets"]} bullets' for role in profile["fixed_roles"]]
    return ", ".join(parts) + ". Exceeding these WILL cause a 2-page resume."


def build_tailor_prompt(company, title, jd, work_dna, base_resumes, profile, base_resume_tags):
    tags = list(base_resume_tags.keys())
    base_content_blocks = "\n\n".join(
        f"--- {tag} ---\n{base_resumes.get(tag, '(not available)')}" for tag in tags
    )
    optional_roles = [r["title"] for r in profile["fixed_roles"] if r.get("optional")]
    optional_note = (
        f"(Optional roles — include only if relevant: {', '.join(optional_roles)})"
        if optional_roles else "(All roles below are included in every tailored resume.)"
    )

    return f"""\
You are an expert resume writer tailoring a one-page resume for {profile['name']}, a {profile['role_noun']}.

=== WORK DNA (who they really are — use this to surface hidden context) ===
{work_dna}

=== BASE RESUME CONTENT (blend the most relevant parts across all of these) ===

{base_content_blocks}

=== TARGET JOB ===
Company  : {company}
Role     : {title}

Choose the best base_tag for this role from: {', '.join(tags)}
Include your choice in the "base_tag" field of the JSON output.

=== JOB DESCRIPTION ===
{jd}

=== TAILORING RULES ===
1. NEVER fabricate — no new roles, employers, dates, or achievements not in Work DNA or base resumes.
2. Blend all base resumes listed above — pick the most relevant bullets and content for this specific JD.
3. Reword existing bullets to surface JD-relevant keywords naturally.
4. Decide which experience roles to include. {optional_note}
5. Decide whether to include a Projects section and which projects (0-2 max) — if the JD explicitly
   asks for proof of work, a portfolio, a GitHub repo, or project links, you MUST include this section
   with at least one real, relevant project — do not treat it as optional in that case.
6. STRICT BULLET LIMITS — do NOT exceed these counts or the resume will overflow to 2 pages:
{_build_bullet_limits(profile)}
7. Keep each bullet concise — target 1.5 lines rendered, never exceed 2.5 lines.
8. Reorder competency categories to put the most JD-relevant first.
9. Fully rewrite the summary to match JD tone, include exact job title from JD.
10. NEVER use en dash (-, \\u2013) inside **bold** markers - use hyphen (-) instead.
11. ORDER MATTERS beyond just competencies: within EVERY list — every role's bullets, projects,
    competencies — put the most important item for THIS JD and THIS role's seniority/vitality
    first, least important last. The resume builder enforces a strict 1-page limit by
    automatically deleting items from the END of these lists if the resume still overflows at
    minimum font size. A weak or irrelevant item placed early can survive a cut that removes
    something better, so ordering is not cosmetic — it decides what actually gets cut.
12. PROTECTED COMPETENCY TERMS — the competency-category cap in rule 6 must never cost you a
    near-universal keyword for this role that Work DNA genuinely supports, just because it only
    appears as one line of prose (e.g. buried in a summary sentence) rather than a named
    competency. A narrative-only mention doesn't survive the full summary/competency rewrite
    every JD triggers. If Work DNA calls out specific terms as protected/must-not-drop, treat
    each as a named Core Competencies term, not something left implied by a bullet or sentence.

=== ATS OPTIMIZATION RULES ===
1. Mirror EXACT JD phrases — use the JD's own words, not synonyms.
2. Top 5 JD keywords must appear at least 2x across summary + bullets combined.
3. Front-load keywords — put the most important JD keywords in the first 2 bullets and summary.
4. No keyword stuffing — keywords must appear in natural sentences.
5. Competencies section must include exact skill terms from the JD as-is.
6. Summary must bridge {profile['name']}'s title to the target role title naturally.
7. Use **bold** markers to highlight key matched terms in bullets.

=== FIXED ELEMENTS (do NOT change these) ===
{_build_fixed_elements(profile)}

=== FIT SCORE FORMULA ===
- 60% weight: % of high-priority JD keywords present in the tailored resume (count matched / total high-priority)
- 40% weight: % of minimum qualifications met by {profile['name']}'s profile
- Hard gaps (missing minimum qualifications) cap the score at 65% maximum
- Output an integer 0-100

=== OUTPUT FORMAT ===
Return ONLY a valid JSON object. No explanation, no markdown, no code fences. The JSON must match this exact schema:

{{
  "base_tag": "{' | '.join(tags)}",
  "summary": "Full summary paragraph. Use **bold** for key terms. No en dash in bold.",
  "experience_header": "PROFESSIONAL EXPERIENCE",
  "experience": {_build_experience_schema_example(profile)},
  "projects_section_label": "KEY PROJECTS AND CAPSTONE WORK",
  "projects": [
    {{"lead": "Project Name", "text": "...", "has_github_link": true}},
    {{"lead": "Second Project", "text": "..."}}
  ],
  "competencies": [
    {{"label": "Most Relevant Category", "items": "item1, item2, item3"}},
    {{"label": "Second Category", "items": "item1, item2"}},
    {{"label": "Third Category", "items": "item1, item2"}}
  ],
  "fit_score": 85,
  "keywords": [
    {{"keyword": "example high-priority keyword", "present": "Yes", "count": 3, "priority": "High"}},
    {{"keyword": "example keyword", "present": "Yes", "count": 2, "priority": "High"}}
  ],
  "gaps": ["Gap area 1", "Gap area 2"],
  "ideas": ["Improvement idea 1", "Improvement idea 2"]
}}

IMPORTANT:
- For any optional role, omit the whole object from "experience" if not relevant to this JD.
- Include "projects" only if there are 1-2 relevant projects — set to [] or omit if none, UNLESS
  the JD explicitly asks for proof-of-work/portfolio/GitHub links (see rule 5 above).
- HARD LIMITS: {_build_hard_limits_line(profile)}
- Keep bullets concise — 1 to 2 lines each. Long bullets are the #1 cause of overflow.
- All bullet text supports **bold** markers for inline bolding.
- All text must be ATS-optimized as described above.
"""


def call_claude_tailor(company, title, jd, work_dna, base_resumes):
    """
    Call claude -p with the full tailoring prompt.
    Tag selection is included in this call — no separate tag pick needed.
    Returns parsed JSON dict or None on failure.
    """
    prompt = build_tailor_prompt(
        company=company, title=title, jd=jd[:4000],
        work_dna=work_dna, base_resumes=base_resumes,
        profile=PROFILE, base_resume_tags=BASE_RESUME_TAGS,
    )

    try:
        with tempfile.NamedTemporaryFile(
            mode='w', suffix='.txt', delete=False, encoding='utf-8'
        ) as f:
            f.write(prompt)
            tmp = f.name

        result = subprocess.run(
            f'type "{tmp}" | claude -p',
            shell=True, capture_output=True, text=True, timeout=300, encoding="utf-8",
        )
        os.unlink(tmp)

        raw = result.stdout.strip()

        # Strip markdown code fences if claude wraps output
        raw = re.sub(r'^```[\w]*\n?', '', raw)
        raw = re.sub(r'\n?```\s*$',   '', raw)
        raw = raw.strip()

        if not raw:
            stderr = (result.stderr or "")[:400]
            print(f"    claude -p returned empty output. stderr: {stderr}")
            return None

        try:
            content = json.loads(raw)
            return content
        except json.JSONDecodeError as e:
            print(f"    JSON parse error: {e} — retrying with fix prompt...")
            # Ask Claude to fix the broken JSON
            fix_prompt = (
                "The following text is almost-valid JSON but has a syntax error. "
                "Return ONLY the corrected JSON — no explanation, no markdown, no code fences.\n\n"
                + raw
            )
            with tempfile.NamedTemporaryFile(
                mode='w', suffix='.txt', delete=False, encoding='utf-8'
            ) as f2:
                f2.write(fix_prompt)
                tmp2 = f2.name
            result2 = subprocess.run(
                f'type "{tmp2}" | claude -p',
                shell=True, capture_output=True, text=True, timeout=120, encoding="utf-8",
            )
            os.unlink(tmp2)
            raw2 = result2.stdout.strip()
            raw2 = re.sub(r'^```[\w]*\n?', '', raw2)
            raw2 = re.sub(r'\n?```\s*$',   '', raw2)
            raw2 = raw2.strip()
            content = json.loads(raw2)
            return content

    except json.JSONDecodeError as e:
        print(f"    JSON still invalid after retry: {e}")
        print(f"    Raw output start: {raw[:500]}")
        return None
    except Exception as e:
        print(f"    claude -p error: {e}")
        return None



# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*62}")
    print(f"  Job Runner — {date.today().isoformat()}")
    print(f"{'='*62}")

    # Guard: Excel lock file
    lock = os.path.join(PIPELINE_DIR, "~$Daily_Jobs_Log.xlsx")
    if os.path.exists(lock):
        print("\n  ERROR: Daily_Jobs_Log.xlsx is open in Excel. Close it and re-run.")
        return

    if not os.path.exists(DAILY_LOG_FILE):
        print("\n  ERROR: Daily_Jobs_Log.xlsx not found. Run job_scraper.py first.")
        return

    # Load base resumes and Work DNA once (reused for every job)
    print("\n  Loading base resumes and Work DNA...")
    base_resumes = load_all_base_resumes()
    work_dna     = load_work_dna()
    loaded_count = sum(1 for v in base_resumes.values() if v)
    print(f"  Loaded {loaded_count}/4 base resumes.  Work DNA: {'OK' if work_dna else 'MISSING'}")

    # Import build_resume
    try:
        from build_resume import auto_build
    except ImportError as e:
        print(f"\n  ERROR: Could not import build_resume: {e}")
        return

    # Open Excel
    wb = openpyxl.load_workbook(DAILY_LOG_FILE)
    ws = wb.active

    ensure_runner_columns(ws)

    # Find shortlisted jobs with no runner status
    jobs = find_shortlisted(ws)
    print(f"\n  Found {len(jobs)} shortlisted job(s) pending tailoring.\n")

    if not jobs:
        print("  Nothing to process.")
        print("  Open Daily_Jobs_Log.xlsx, set Status = 'Shortlisted' on jobs you like, then re-run.")
        wb.save(DAILY_LOG_FILE)
        return

    succeeded = 0
    failed    = 0

    for row_num, company, title, jd, url in jobs:
        print(f"\n  {'─'*55}")
        print(f"  Job : {company} / {title}")
        if url:
            print(f"  URL : {url[:80]}")

        # Skip if no JD
        if not jd or jd.startswith("No JD"):
            print("  Skipped — no JD available.")
            print("  Paste the JD into the 'Full JD' cell, then clear 'Runner Status' to retry.")
            set_runner_status(ws, row_num, "", "Skipped-NoJD")
            wb.save(DAILY_LOG_FILE)
            continue

        today        = date.today().isoformat()
        company_safe = sanitize(company)

        # Step 1: Tailor resume — tag selection + full content in one call
        print("  [1/2] Tailoring resume via claude -p (may take ~90s)...")
        content = call_claude_tailor(company, title, jd, work_dna, base_resumes)

        if not content:
            print("  FAILED — claude -p returned no usable content.")
            set_runner_status(ws, row_num, "Unknown", "Failed")
            wb.save(DAILY_LOG_FILE)
            failed += 1
            continue

        default_tag = next(iter(BASE_RESUME_TAGS))
        final_tag   = content.get("base_tag", default_tag)
        print(f"        → {final_tag}")

        # Create output folder using tag from response
        folder_name = f"{company_safe}_{final_tag}_{today}"
        out_dir     = os.path.join(JOB_APPS_DIR, folder_name)
        os.makedirs(out_dir, exist_ok=True)

        # Step 2: Build resume
        name_prefix = sanitize(PROFILE["name"], maxlen=30)
        out_name = f"{name_prefix}_{final_tag}_{company_safe}_{today}"
        print(f"  [2/2] Building resume (auto-scale binary search)...")
        try:
            docx_path, pdf_path, fs = auto_build(content, out_dir, out_name)
            print(f"        Font: {fs:.2f}pt")
        except Exception as e:
            print(f"  FAILED — build_resume error: {e}")
            set_runner_status(ws, row_num, final_tag, "Failed")
            wb.save(DAILY_LOG_FILE)
            failed += 1
            continue

        # Mark done — ATS analysis, fit score, master tracker handled by job_analysis.py
        set_runner_status(ws, row_num, final_tag, "Done")
        wb.save(DAILY_LOG_FILE)
        succeeded += 1

        print(f"\n  ✓  Resume done → {out_dir}")
        print(f"     Run job_analysis.py to generate ATS report and fit score.")

    print(f"\n{'='*62}")
    print(f"  Done.  Succeeded: {succeeded}  |  Failed: {failed}")
    print(f"  Output dir : {JOB_APPS_DIR}")
    print(f"  Log        : {DAILY_LOG_FILE}")
    print(f"{'='*62}\n")


if __name__ == "__main__":
    main()
