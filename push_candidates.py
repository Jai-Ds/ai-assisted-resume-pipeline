# -*- coding: utf-8 -*-
"""
push_candidates.py — the real push step, compared against the local decisions log, never
the live dashboard.

Why: the dashboard has a "Clear today's list" button that archives its whole `jobs`
collection. The dedup check must not depend on that database still looking the way it did
yesterday — so this script compares new candidates against Application_Decisions_Log.xlsx
(see decisions_log.py), a separate local file nothing on the dashboard can touch.

The dashboard database can only be written to by a live Claude Code tool call (ArtifactData)
— no local script, including this one, can reach it directly. So the real workflow is:

  1. (live session) Fetch the dashboard's current `jobs` collection, merge into one JSON list.
  2. (live session) Call decisions_log.sync_from_dashboard_dump(...) with that JSON, to pull
     any new Applied/Skipped decisions into the log. This is the ONLY step where the dashboard
     touches the log — after this, the log stands on its own.
  3. Run this script: python push_candidates.py --candidates candidates.json
     --decisions-log "Application_Decisions_Log.xlsx" --out result.json
  4. (live session) Batch-write result.json's "to_push" list to the dashboard via ArtifactData.
  5. (live session) Call decisions_log.record_pushed(...) with that same "to_push" list, so
     these jobs are in the log (blank status) even before the user decides on them — this is
     what makes the exact-URL idempotency check work without ever re-reading the dashboard.

candidates.json: a JSON list of {"title", "company", "location", "apply_url", "platform",
    "date", "jd_text", "note"} — jd_text is usually "" at scrape time.

result.json: {"to_push": [...candidates with a "doc_id" field added...],
              "dropped": [...candidates with a "_dedup_reason" field added...]}
"""

import argparse
import hashlib
import json
import sys

sys.stdout.reconfigure(encoding="utf-8")  # Windows cp1252 console otherwise garbles the em dash below

from dedup_check import filter_new_jobs
from decisions_log import load_decisions


def compute_doc_id(apply_url):
    return "job_" + hashlib.sha1(apply_url.encode("utf-8")).hexdigest()[:16]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidates", required=True, help="Path to JSON list of newly-scraped candidate jobs")
    ap.add_argument("--decisions-log", required=True, help="Path to Application_Decisions_Log.xlsx")
    ap.add_argument("--out", required=True, help="Path to write the filtered result JSON")
    ap.add_argument("--window-days", type=int, default=30)
    args = ap.parse_args()

    with open(args.candidates, encoding="utf-8") as f:
        candidates = json.load(f)

    decisions = load_decisions(args.decisions_log)

    for c in candidates:
        if not c.get("apply_url"):
            raise ValueError(f"Candidate missing apply_url, cannot compute doc_id: {c}")
        c["doc_id"] = compute_doc_id(c["apply_url"])

    # Idempotency guard, separate from the 30-day decided-dedup rule: never re-push the exact
    # same apply_url twice, decided or not — decisions_log.record_pushed() logs every pushed
    # job (even undecided ones), so this check works without ever reading the dashboard.
    logged_doc_ids = {d["doc_id"] for d in decisions if d.get("doc_id")}
    already_present = [c for c in candidates if c["doc_id"] in logged_doc_ids]
    candidates = [c for c in candidates if c["doc_id"] not in logged_doc_ids]
    for c in already_present:
        c["_dedup_reason"] = "exact same listing URL already logged (pushed before, decided or not)"

    keep, dropped = filter_new_jobs(candidates, decisions, window_days=args.window_days)
    dropped = already_present + dropped

    result = {"to_push": keep, "dropped": dropped}
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print(f"{len(keep)} to push, {len(dropped)} dropped (of {len(keep) + len(dropped)} candidates)")
    for d in dropped:
        print(f"  DROPPED: {d.get('title')!r} @ {d.get('company')!r} — {d.get('_dedup_reason')}")


if __name__ == "__main__":
    main()
