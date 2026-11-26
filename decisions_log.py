# -*- coding: utf-8 -*-
"""
decisions_log.py — a local, persistent record of job decisions (Applied/Skipped/blank),
independent of the live dashboard database.

Why this exists: the dashboard has a "Clear today's list" button that archives every job in
its own database. The dedup check must survive that — it should never depend on reading the
live, mutable dashboard at push time. This Excel file is the single source of truth for
dedup going forward; the dashboard is only ever a one-way SOURCE that gets synced INTO this
file, never the thing compared against directly.

File: Application_Decisions_Log.xlsx, one row per job (keyed by Doc ID — the same stable
sha1-of-apply_url hash used elsewhere in this pipeline).

Columns: Doc ID | Title | Company | Platform | Apply URL | Status | Status Date | First Seen

- Status is "Applied", "Skipped", or blank (never decided).
- Status Date is a plain YYYY-MM-DD string, set the day a status is recorded.
- First Seen is the date this job was first pushed to the dashboard, for reference only —
  not used by the dedup check itself (only Status / Status Date matter there).

Two local operations, both plain file I/O, no dashboard access:
  sync_from_dashboard_dump(dump_path, log_path)
      Call this with a fresh JSON dump of the dashboard's `jobs` collection (from
      ArtifactData `list`/`query` with `out_dir`, merged into one JSON list — see
      push_candidates.py's companion step) to pull the latest Applied/Skipped decisions
      into the log. This is the ONLY point where dashboard data touches this file.
  record_pushed(candidates, log_path)
      Call this right after a successful dashboard push, so newly-pushed jobs are in the
      log (with a blank Status) even before the user decides on them.

load_decisions(log_path) -> list[dict] shaped like dedup_check.py expects
  ({"title", "company", "user_action", "user_action_ts"}) — this is what push_candidates.py
  actually compares new candidates against.
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime, date

sys.stdout.reconfigure(encoding="utf-8")

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

HEADERS = ["Doc ID", "Title", "Company", "Platform", "Apply URL", "Status", "Status Date", "First Seen"]
COL_WIDTHS = [20, 45, 30, 12, 60, 12, 14, 14]

_STATUS_TO_ACTION = {"applied": "apply", "skipped": "skip"}
_ACTION_TO_STATUS = {"apply": "Applied", "skip": "Skipped"}


def _new_workbook():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Decisions"
    for col, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="2E75B6")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    return wb, ws


def _open_or_create(log_path):
    if os.path.exists(log_path):
        wb = openpyxl.load_workbook(log_path)
        ws = wb.active
        return wb, ws
    return _new_workbook()


def _row_index_by_doc_id(ws):
    idx = {}
    for r in range(2, ws.max_row + 1):
        doc_id = ws.cell(r, 1).value
        if doc_id:
            idx[doc_id] = r
    return idx


def _date_from_ms(ms):
    try:
        return datetime.utcfromtimestamp(ms / 1000).strftime("%Y-%m-%d")
    except Exception:
        return ""


def sync_from_dashboard_dump(dashboard_jobs, log_path):
    """
    dashboard_jobs: list of dicts from the dashboard's `jobs` collection (must include
    "id"/"doc_id", "title", "company", "platform", "apply_url", "user_action", "user_action_ts").
    Upserts Status/Status Date for any job with a decided user_action. Never removes rows —
    this file only grows, by design (it must outlive dashboard clears).
    """
    wb, ws = _open_or_create(log_path)
    row_by_id = _row_index_by_doc_id(ws)
    today = date.today().isoformat()
    updated = 0

    for j in dashboard_jobs:
        doc_id = j.get("id") or j.get("doc_id")
        if not doc_id:
            continue
        action = j.get("user_action")
        status = _ACTION_TO_STATUS.get(action, "")
        status_date = _date_from_ms(j.get("user_action_ts")) if status else ""

        if doc_id in row_by_id:
            r = row_by_id[doc_id]
            if status and ws.cell(r, 6).value != status:
                ws.cell(r, 6).value = status
                ws.cell(r, 7).value = status_date or today
                updated += 1
            # keep title/company/platform/url fresh in case they were blank before
            if not ws.cell(r, 2).value:
                ws.cell(r, 2).value = j.get("title", "")
            if not ws.cell(r, 3).value:
                ws.cell(r, 3).value = j.get("company", "")
            if not ws.cell(r, 4).value:
                ws.cell(r, 4).value = j.get("platform", "")
            if not ws.cell(r, 5).value:
                ws.cell(r, 5).value = j.get("apply_url", "")
        else:
            r = ws.max_row + 1
            ws.cell(r, 1).value = doc_id
            ws.cell(r, 2).value = j.get("title", "")
            ws.cell(r, 3).value = j.get("company", "")
            ws.cell(r, 4).value = j.get("platform", "")
            ws.cell(r, 5).value = j.get("apply_url", "")
            ws.cell(r, 6).value = status
            ws.cell(r, 7).value = status_date
            ws.cell(r, 8).value = j.get("date", today)
            row_by_id[doc_id] = r
            if status:
                updated += 1

    wb.save(log_path)
    return updated


def record_pushed(candidates, log_path):
    """candidates: list of dicts with 'doc_id', 'title', 'company', 'platform', 'apply_url'
    (the 'to_push' list push_candidates.py just wrote to the dashboard) — adds a blank-status
    row for each one not already in the log, so the idempotency check sees them next time."""
    wb, ws = _open_or_create(log_path)
    row_by_id = _row_index_by_doc_id(ws)
    today = date.today().isoformat()
    added = 0

    for c in candidates:
        doc_id = c.get("doc_id")
        if not doc_id or doc_id in row_by_id:
            continue
        r = ws.max_row + 1
        ws.cell(r, 1).value = doc_id
        ws.cell(r, 2).value = c.get("title", "")
        ws.cell(r, 3).value = c.get("company", "")
        ws.cell(r, 4).value = c.get("platform", "")
        ws.cell(r, 5).value = c.get("apply_url", "")
        ws.cell(r, 6).value = ""
        ws.cell(r, 7).value = ""
        ws.cell(r, 8).value = today
        row_by_id[doc_id] = r
        added += 1

    wb.save(log_path)
    return added


def load_decisions(log_path):
    """Returns a list of dicts shaped for dedup_check.py: {"doc_id", "title", "company",
    "apply_url", "user_action", "user_action_ts"}. If the file doesn't exist yet, returns []
    (nothing decided yet — every candidate will pass the dedup check, which is correct)."""
    if not os.path.exists(log_path):
        return []
    wb = openpyxl.load_workbook(log_path, read_only=True)
    ws = wb.active
    out = []
    for r in range(2, ws.max_row + 1):
        doc_id = ws.cell(r, 1).value
        if not doc_id:
            continue
        status = (ws.cell(r, 6).value or "").strip().lower()
        status_date = ws.cell(r, 7).value
        ts = None
        if status_date:
            try:
                ts = int(datetime.strptime(str(status_date)[:10], "%Y-%m-%d").timestamp() * 1000)
            except Exception:
                ts = None
        out.append({
            "doc_id": doc_id,
            "title": ws.cell(r, 2).value or "",
            "company": ws.cell(r, 3).value or "",
            "apply_url": ws.cell(r, 5).value or "",
            "user_action": _STATUS_TO_ACTION.get(status),
            "user_action_ts": ts,
        })
    return out


def _cli():
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_sync = sub.add_parser("sync", help="Pull Applied/Skipped decisions from a dashboard jobs-collection dump into the log")
    p_sync.add_argument("--dump", required=True, help="Path to a JSON list of dashboard `jobs` collection docs")
    p_sync.add_argument("--log", required=True, help="Path to Application_Decisions_Log.xlsx")

    p_rec = sub.add_parser("record-pushed", help="Log newly-pushed jobs (blank status) so future idempotency checks see them")
    p_rec.add_argument("--pushed", required=True, help="Path to a JSON list of jobs just pushed (needs doc_id/title/company/platform/apply_url)")
    p_rec.add_argument("--log", required=True, help="Path to Application_Decisions_Log.xlsx")

    args = ap.parse_args()

    if args.cmd == "sync":
        with open(args.dump, encoding="utf-8") as f:
            dump = json.load(f)
        updated = sync_from_dashboard_dump(dump, args.log)
        print(f"Synced {updated} newly-decided row(s) into {args.log}")
    elif args.cmd == "record-pushed":
        with open(args.pushed, encoding="utf-8") as f:
            pushed = json.load(f)
        added = record_pushed(pushed, args.log)
        print(f"Logged {added} newly-pushed job(s) into {args.log}")


if __name__ == "__main__":
    _cli()
