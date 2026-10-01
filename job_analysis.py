# -*- coding: utf-8 -*-
"""
job_analysis.py — Resume ATS Analysis Runner

Runs AFTER job_runner.py has built resumes.
Finds rows in Daily_Jobs_Log.xlsx where:
  - Runner Status = "Done"
  - Analysis Status (col 13) is blank

For each qualifying job:
  1. Locates the company folder in Job Applications/
  2. Reads the built DOCX resume text
  3. Calls claude -p with resume + JD → ATS analysis JSON
  4. Creates ATS_Analysis_[Company].xlsx in the company folder
  5. Updates Job_Applications_Master.xlsx
  6. Writes Fit Score (col 12) and Analysis Status (col 13) back to Excel

Run separately after job_runner.py, or schedule after it.
Designed to grow: add more analysis blocks below the existing ones.
"""

import os
import sys
import re
import json
import glob
import textwrap
import subprocess
import tempfile
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from docx import Document

sys.stdout.reconfigure(encoding="utf-8")


# ── Config ────────────────────────────────────────────────────────────────────
# Paths come from config.json — copy config.example.json to get started.

from config import PATHS

PIPELINE_DIR   = PATHS["pipeline_dir"]
DAILY_LOG_FILE = os.path.join(PIPELINE_DIR, "Daily_Jobs_Log.xlsx")
JOB_APPS_DIR   = PATHS["job_apps_dir"]

sys.path.insert(0, JOB_APPS_DIR)

# Column indices (1-based)
COL_COMPANY     = 2
COL_TITLE       = 3
COL_URL         = 5
COL_JD          = 6
COL_STATUS      = 7
COL_BASE_RESUME = 10
COL_RUNNER_ST   = 11
COL_FIT_SCORE   = 12
COL_ANALYSIS_ST = 13

HEADER_ROW = 2

MID_BLUE = "2E75B6"
WHITE    = "FFFFFF"


# ── Excel helpers ─────────────────────────────────────────────────────────────

def ensure_analysis_column(ws):
    """Add Analysis Status (col 13) header if not present."""
    cell = ws.cell(HEADER_ROW, COL_ANALYSIS_ST)
    if cell.value != "Analysis Status":
        cell.value     = "Analysis Status"
        cell.font      = Font(bold=True, color=WHITE, size=10)
        cell.fill      = PatternFill("solid", fgColor=MID_BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(COL_ANALYSIS_ST)].width = 15


def find_pending_analysis(ws):
    """
    Return list of (row_num, company, title, jd, url, base_tag) where:
      - Runner Status = "Done"
      - Analysis Status is blank
    """
    results = []
    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=False):
        runner_st   = row[COL_RUNNER_ST   - 1].value
        analysis_st = row[COL_ANALYSIS_ST - 1].value
        if runner_st == "Done" and not analysis_st:
            company  = row[COL_COMPANY     - 1].value or ""
            title    = row[COL_TITLE       - 1].value or ""
            jd       = row[COL_JD          - 1].value or ""
            url      = row[COL_URL         - 1].value or ""
            base_tag = row[COL_BASE_RESUME - 1].value or ""
            results.append((row[0].row, company, title, jd, url, base_tag))
    return results


def set_analysis_result(ws, row_num, fit_score, status):
    """Write Fit Score and Analysis Status into the Excel row."""
    if fit_score is not None:
        cell           = ws.cell(row_num, COL_FIT_SCORE)
        cell.value     = fit_score
        cell.font      = Font(size=9, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if fit_score >= 80:
            cell.fill = PatternFill("solid", fgColor="C6EFCE")
        elif fit_score >= 60:
            cell.fill = PatternFill("solid", fgColor="FFEB9C")
        else:
            cell.fill = PatternFill("solid", fgColor="FFC7CE")

    cell           = ws.cell(row_num, COL_ANALYSIS_ST)
    cell.value     = status
    cell.font      = Font(size=9)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ── Folder + DOCX finder ──────────────────────────────────────────────────────

def sanitize(text, maxlen=25):
    return re.sub(r'[^\w]', '', (text or "").replace(' ', '_'))[:maxlen]


def find_company_folder(company, base_tag):
    """
    Search JOB_APPS_DIR for the most recent folder matching
    {sanitize(company)}_{base_tag}_* pattern.
    Returns folder path or None.
    """
    company_safe = sanitize(company)
    prefix       = f"{company_safe}_{base_tag}_"
    matches      = [
        os.path.join(JOB_APPS_DIR, d)
        for d in os.listdir(JOB_APPS_DIR)
        if os.path.isdir(os.path.join(JOB_APPS_DIR, d)) and d.startswith(prefix)
    ]
    if not matches:
        return None
    # Most recently modified folder wins
    return max(matches, key=os.path.getmtime)


def read_docx_text(folder_path):
    """
    Find the tailored resume .docx in the folder and extract its text.
    Each company folder holds exactly one resume docx, so any *.docx match works.
    Returns text string or empty string.
    """
    pattern = os.path.join(folder_path, "*.docx")
    matches = glob.glob(pattern)
    if not matches:
        return ""
    try:
        doc   = Document(matches[0])
        lines = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(lines)
    except Exception as e:
        print(f"    WARNING: Could not read DOCX: {e}")
        return ""


# ── Claude ATS analysis call ──────────────────────────────────────────────────

ATS_PROMPT = """\
You are an ATS analyst. Evaluate the resume below against the job description and return a JSON analysis.

COMPANY : {company}
ROLE    : {title}

RESUME TEXT:
{resume_text}

JOB DESCRIPTION:
{jd}

SCORING FORMULA:
- 60% weight: % of high-priority JD keywords present in the resume (count matched / total high-priority)
- 40% weight: % of minimum qualifications met by the candidate
- Hard gaps (missing minimum qualifications) cap the total score at 65% maximum
- Output an integer fit_score 0-100

INSTRUCTIONS:
1. Identify the top JD keywords (10-15 total: mix of High/Med/Low priority).
2. Check each keyword in the resume text (case-insensitive). Count occurrences.
3. Identify real gaps — skills or qualifications the JD requires that the resume lacks.
4. Suggest practical improvement ideas (rewordings, additions from real experience only).
5. Return ONLY valid JSON. No explanation, no markdown, no code fences.

JSON schema:
{{
  "fit_score": 82,
  "keywords": [
    {{"keyword": "supply chain planning", "present": "Yes", "count": 3, "priority": "High"}},
    {{"keyword": "stakeholder management", "present": "No",  "count": 0, "priority": "High"}}
  ],
  "gaps": [
    "Gap description 1",
    "Gap description 2"
  ],
  "ideas": [
    "Improvement idea 1",
    "Improvement idea 2"
  ]
}}
"""


def call_claude_analysis(company, title, resume_text, jd):
    """
    Call claude -p to analyse resume vs JD.
    Returns parsed JSON dict or None on failure.
    """
    prompt = ATS_PROMPT.format(
        company      = company,
        title        = title,
        # NOTE: a 1-page resume's extracted text runs ~4000-4500 chars, and
        # Core Competencies — the section most keyword-matches actually live
        # in — is always last. A 3000-char cutoff here silently excluded it
        # entirely (confirmed 2026-09-27: real terms at chars 3899/4131 were
        # scored "No" purely because they were never sent to the model).
        resume_text  = resume_text[:6000],
        jd           = jd[:6000],
    )

    try:
        with tempfile.NamedTemporaryFile(
            mode='w', suffix='.txt', delete=False, encoding='utf-8'
        ) as f:
            f.write(prompt)
            tmp = f.name

        result = subprocess.run(
            f'type "{tmp}" | claude -p',
            shell=True, capture_output=True, text=True, timeout=120, encoding="utf-8",
        )
        os.unlink(tmp)

        raw = result.stdout.strip()
        raw = re.sub(r'^```[\w]*\n?', '', raw)
        raw = re.sub(r'\n?```\s*$',   '', raw)
        raw = raw.strip()

        if not raw:
            print(f"    claude -p returned empty. stderr: {(result.stderr or '')[:300]}")
            return None

        return json.loads(raw)

    except json.JSONDecodeError as e:
        print(f"    JSON parse error: {e}")
        print(f"    Raw output: {raw[:400]}")
        return None
    except Exception as e:
        print(f"    claude -p error: {e}")
        return None


# ── Analysis blocks ───────────────────────────────────────────────────────────
# Add future blocks (cover letter, interview prep, etc.) as new functions here.

def block_ats_excel(folder_path, company, title, analysis):
    """Block 1: Create ATS_Analysis_[Company].xlsx in the company folder."""
    try:
        from job_tracker import create_ats_sheet
        ats_path = create_ats_sheet(
            folder_path = folder_path,
            company     = company,
            role_title  = title,
            keywords    = analysis.get("keywords", []),
            fit_score   = analysis.get("fit_score", 0),
            gaps        = analysis.get("gaps",     []),
            ideas       = analysis.get("ideas",    []),
        )
        print(f"    ATS Excel  : {os.path.basename(ats_path)}")
        return True
    except Exception as e:
        print(f"    ATS Excel failed: {e}")
        return False


def block_master_tracker(company, title, fit_score, folder_path, url):
    """Block 2: Update Job_Applications_Master.xlsx."""
    try:
        from job_tracker import update_master
        update_master(
            company     = company,
            role_title  = title,
            fit_score   = fit_score,
            folder_path = folder_path,
            job_url     = url,
            date_str    = date.today().isoformat(),
        )
        print(f"    Master     : updated")
        return True
    except Exception as e:
        print(f"    Master tracker failed: {e}")
        return False


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*62}")
    print(f"  Job Analysis — {date.today().isoformat()}")
    print(f"{'='*62}")

    # Guard: Excel lock file
    lock = os.path.join(PIPELINE_DIR, "~$Daily_Jobs_Log.xlsx")
    if os.path.exists(lock):
        print("\n  ERROR: Daily_Jobs_Log.xlsx is open in Excel. Close it and re-run.")
        return

    if not os.path.exists(DAILY_LOG_FILE):
        print("\n  ERROR: Daily_Jobs_Log.xlsx not found.")
        return

    wb = openpyxl.load_workbook(DAILY_LOG_FILE)
    ws = wb.active

    ensure_analysis_column(ws)

    jobs = find_pending_analysis(ws)
    print(f"\n  Found {len(jobs)} job(s) pending analysis.\n")

    if not jobs:
        print("  Nothing to analyse.")
        print("  Run job_runner.py first to build resumes, then re-run this script.")
        wb.save(DAILY_LOG_FILE)
        return

    succeeded = 0
    failed    = 0

    for row_num, company, title, jd, url, base_tag in jobs:
        print(f"\n  {'─'*55}")
        print(f"  Job : {company} / {title}")

        # Skip if no JD
        if not jd or jd.startswith("No JD"):
            print("  Skipped — no JD in Excel.")
            set_analysis_result(ws, row_num, None, "Skipped-NoJD")
            wb.save(DAILY_LOG_FILE)
            continue

        # Step 1: Find company folder
        folder_path = find_company_folder(company, base_tag)
        if not folder_path:
            print(f"  Skipped — company folder not found for: {sanitize(company)}_{base_tag}_*")
            set_analysis_result(ws, row_num, None, "Skipped-NoFolder")
            wb.save(DAILY_LOG_FILE)
            continue
        print(f"  Folder : {os.path.basename(folder_path)}")

        # Step 2: Read resume DOCX
        resume_text = read_docx_text(folder_path)
        if not resume_text:
            print("  Skipped — no DOCX found in folder.")
            set_analysis_result(ws, row_num, None, "Skipped-NoFolder")
            wb.save(DAILY_LOG_FILE)
            continue

        # Step 3: Call claude -p for ATS analysis
        print("  [1/3] Running ATS analysis via claude -p...")
        analysis = call_claude_analysis(company, title, resume_text, jd)

        if not analysis:
            print("  FAILED — claude -p returned no usable analysis.")
            set_analysis_result(ws, row_num, None, "Failed")
            wb.save(DAILY_LOG_FILE)
            failed += 1
            continue

        fit_score = analysis.get("fit_score", 0)
        keywords  = analysis.get("keywords", [])
        gaps      = analysis.get("gaps",     [])

        # Step 4: ATS Excel
        print("  [2/3] Creating ATS Excel...")
        block_ats_excel(folder_path, company, title, analysis)

        # Step 5: Master tracker
        print("  [3/3] Updating master tracker...")
        block_master_tracker(company, title, fit_score, folder_path, url)

        # Write back to Excel
        set_analysis_result(ws, row_num, fit_score, "Done")
        wb.save(DAILY_LOG_FILE)
        succeeded += 1

        # Terminal summary
        print(f"\n  ✓  Analysis done")
        print(f"     Fit Score : {fit_score}%")
        high = [k for k in keywords if k.get("priority") == "High"]
        matched = sum(1 for k in high if k.get("present") == "Yes")
        print(f"     Keywords  : {matched}/{len(high)} high-priority matched")
        for k in high:
            mark = "✓" if k.get("present") == "Yes" else "✗"
            print(f"       {mark} {k['keyword']} (×{k.get('count', 0)})")
        if gaps:
            print(f"     Gaps      :")
            for g in gaps:
                print(f"       • {g}")

        # ── Future analysis blocks go here ─────────────────────────────────
        # block_cover_letter(folder_path, company, title, jd, resume_text)
        # block_interview_prep(folder_path, company, title, jd)
        # block_company_research(folder_path, company)
        # ───────────────────────────────────────────────────────────────────

    print(f"\n{'='*62}")
    print(f"  Done.  Succeeded: {succeeded}  |  Failed: {failed}")
    print(f"  Output dir : {JOB_APPS_DIR}")
    print(f"  Log        : {DAILY_LOG_FILE}")
    print(f"{'='*62}\n")


if __name__ == "__main__":
    main()
