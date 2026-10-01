# -*- coding: utf-8 -*-
"""
job_tracker.py — Resume application tracker utility
- Creates per-company folder with DOCX, PDF, build script, ATS analysis
- Maintains master Job_Applications_Master.xlsx
"""

import os, shutil
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# config.json lives in Job Pipeline/ — callers (job_runner.py, job_analysis.py)
# already put that directory on sys.path before importing this module.
from config import PATHS

JOB_APPS_DIR = PATHS["job_apps_dir"]
MASTER_FILE  = os.path.join(JOB_APPS_DIR, "Job_Applications_Master.xlsx")


# ── Helpers ────────────────────────────────────────────────────────────────────

def _hfill(color):  return PatternFill("solid", fgColor=color)
def _hfont(**kw):   return Font(**kw)
def _center():      return Alignment(horizontal="center", vertical="center", wrap_text=True)
def _left():        return Alignment(horizontal="left",   vertical="center", wrap_text=True)


# ── Step 1: Create company folder & move files ─────────────────────────────────

def create_company_folder(company, role_tag, date_str):
    """Returns path to created folder like .../Spendflo_AI_PM_2026-03-28/"""
    folder_name = f"{company}_{role_tag}_{date_str}"
    folder_path = os.path.join(JOB_APPS_DIR, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def move_resume_files(folder_path, out_name, build_script):
    """
    Move DOCX + PDF into company folder.
    Copy (not move) the build script so it stays runnable from root too.
    out_name: e.g. 'Jayandran_PM_AI_Spendflo_2026-03-28'
    build_script: e.g. 'build_spendflo.py'
    """
    moved = []
    for ext in [".docx", ".pdf"]:
        src = os.path.join(JOB_APPS_DIR, out_name + ext)
        dst = os.path.join(folder_path, out_name + ext)
        if os.path.exists(src):
            shutil.copy2(src, dst)   # copy first
            try:
                os.remove(src)       # then delete source
            except PermissionError:
                print(f"  NOTE: {ext[1:].upper()} is open — copied but source not deleted. Close the file and delete manually: {src}")
            moved.append(dst)
        elif os.path.exists(dst):
            moved.append(dst)  # already moved

    if build_script:
        src_py = os.path.join(JOB_APPS_DIR, build_script)
        dst_py = os.path.join(folder_path, build_script)
        if os.path.exists(src_py):
            shutil.copy2(src_py, dst_py)

    return moved


# ── Step 2: ATS Analysis Excel ─────────────────────────────────────────────────

def create_ats_sheet(folder_path, company, role_title, keywords, fit_score, gaps, ideas):
    """
    keywords: list of dicts — {keyword, present (Yes/No), count, priority (High/Med/Low)}
    gaps:     list of strings
    ideas:    list of strings
    """
    wb = openpyxl.Workbook()

    # ── Sheet 1: Keyword Match ──────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Keyword Match"

    BLUE   = "2F4F8F"
    GREEN  = "C6EFCE"
    RED    = "FFC7CE"
    YELLOW = "FFEB9C"

    # Title
    ws1["A1"] = f"ATS Keyword Analysis  |  {company}  |  {role_title}"
    ws1.merge_cells("A1:D1")
    ws1["A1"].font  = Font(bold=True, size=13)
    ws1["A1"].alignment = _center()
    ws1.row_dimensions[1].height = 24

    ws1.append([])  # blank row

    # Header row
    for col, h in enumerate(["JD Keyword", "In Resume", "Frequency", "Priority"], 1):
        c = ws1.cell(row=3, column=col, value=h)
        c.fill      = _hfill(BLUE)
        c.font      = Font(bold=True, color="FFFFFF", size=11)
        c.alignment = _center()

    ws1.row_dimensions[3].height = 20

    # Data rows
    for kw in keywords:
        ws1.append([kw["keyword"], kw["present"], kw["count"], kw["priority"]])
        r = ws1.max_row
        ws1.cell(r, 1).alignment = _left()
        b_cell = ws1.cell(r, 2)
        b_cell.alignment = _center()
        b_cell.fill = _hfill(GREEN) if kw["present"] == "Yes" else _hfill(RED)
        ws1.cell(r, 3).alignment = _center()
        ws1.cell(r, 4).alignment = _center()

    ws1.column_dimensions["A"].width = 42
    ws1.column_dimensions["B"].width = 13
    ws1.column_dimensions["C"].width = 12
    ws1.column_dimensions["D"].width = 12

    # ── Sheet 2: Fit Report ─────────────────────────────────────────────────
    ws2 = wb.create_sheet("Fit Report")

    score_color = GREEN if fit_score >= 80 else (YELLOW if fit_score >= 60 else RED)

    ws2["A1"] = "Fit Report"
    ws2.merge_cells("A1:B1")
    ws2["A1"].font      = Font(bold=True, size=14)
    ws2["A1"].alignment = _center()
    ws2.row_dimensions[1].height = 26

    ws2.append([])

    ws2["A3"] = "Fit Score"
    ws2["B3"] = f"{fit_score}%"
    ws2["A3"].font      = Font(bold=True, size=13)
    ws2["B3"].font      = Font(bold=True, size=13)
    ws2["B3"].fill      = _hfill(score_color)
    ws2["B3"].alignment = _center()
    ws2.row_dimensions[3].height = 22

    ws2.append([])

    # Verdict
    verdict = "Strong Fit" if fit_score >= 80 else ("Moderate Fit — Review Gaps" if fit_score >= 60 else "Stretch Role — Significant Gaps")
    ws2["A5"] = "Verdict"
    ws2["B5"] = verdict
    ws2["A5"].font = Font(bold=True)
    ws2["B5"].font = Font(italic=True)
    ws2.row_dimensions[5].height = 18

    ws2.append([])

    # Gap Areas
    ws2["A7"] = "Gap Areas"
    ws2["A7"].font = Font(bold=True, color="C0392B", size=11)
    for i, gap in enumerate(gaps, 8):
        ws2.cell(row=i, column=1, value="")
        c = ws2.cell(row=i, column=2, value=f"• {gap}")
        c.alignment = _left()
        c.fill = _hfill("FFF0F0")
        ws2.row_dimensions[i].height = 30

    next_row = 8 + len(gaps) + 1

    # Improvement Ideas
    ws2.cell(row=next_row, column=1, value="Improvement Ideas")
    ws2.cell(row=next_row, column=1).font = Font(bold=True, color="1F618D", size=11)
    for j, idea in enumerate(ideas, next_row + 1):
        c = ws2.cell(row=j, column=2, value=f"• {idea}")
        c.alignment = _left()
        c.fill = _hfill("EBF5FB")
        ws2.row_dimensions[j].height = 36

    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 80

    out_path = os.path.join(folder_path, f"ATS_Analysis_{company}.xlsx")
    wb.save(out_path)
    return out_path


# ── Step 3: Update master tracker ─────────────────────────────────────────────

def update_master(company, role_title, fit_score, folder_path, job_url="", date_str=None):
    """Append a row to Job_Applications_Master.xlsx (creates it if missing)."""
    if date_str is None:
        date_str = datetime.now().strftime("%Y-%m-%d")

    added_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    BLUE   = "2F4F8F"
    GREEN  = "C6EFCE"
    YELLOW = "FFEB9C"
    RED    = "FFC7CE"

    headers = [
        "Company", "Role Title", "Date Added", "Fit Score (%)",
        "Resume Folder Path", "Job Posting URL",
        "Status", "Date Applied", "Interview Round", "Outcome", "Notes"
    ]
    col_widths = [20, 45, 18, 14, 65, 55, 15, 15, 16, 15, 35]

    if os.path.exists(MASTER_FILE):
        wb = openpyxl.load_workbook(MASTER_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Job Applications"
        for col, (h, w) in enumerate(zip(headers, col_widths), 1):
            c = ws.cell(row=1, column=col, value=h)
            c.fill      = _hfill(BLUE)
            c.font      = Font(bold=True, color="FFFFFF", size=11)
            c.alignment = _center()
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"

    score_color = GREEN if fit_score >= 80 else (YELLOW if fit_score >= 60 else RED)

    row_data = [company, role_title, added_at, fit_score,
                folder_path, job_url, "", "", "", "", ""]
    ws.append(row_data)
    r = ws.max_row
    ws.row_dimensions[r].height = 18

    score_cell = ws.cell(r, 4)
    score_cell.fill      = _hfill(score_color)
    score_cell.font      = Font(bold=True)
    score_cell.alignment = _center()

    for col in range(1, len(headers) + 1):
        ws.cell(r, col).alignment = _left()

    wb.save(MASTER_FILE)
    return MASTER_FILE


# ── Main entry: call this after every build ────────────────────────────────────

def finalize(company, role_tag, role_title, date_str, out_name, build_script,
             keywords, fit_score, gaps, ideas, job_url=""):
    """
    Full pipeline: folder → move files → ATS sheet → master update.
    Returns dict with paths.
    """
    folder  = create_company_folder(company, role_tag, date_str)
    moved   = move_resume_files(folder, out_name, build_script)
    ats     = create_ats_sheet(folder, company, role_title, keywords, fit_score, gaps, ideas)
    master  = update_master(company, role_title, fit_score, folder, job_url, date_str)

    print(f"\n  Folder : {folder}")
    print(f"  Files  : {len(moved)} resume file(s) moved")
    print(f"  ATS    : {os.path.basename(ats)}")
    print(f"  Master : {os.path.basename(master)}")
    return {"folder": folder, "ats": ats, "master": master}
