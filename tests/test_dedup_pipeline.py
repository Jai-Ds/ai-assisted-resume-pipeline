# -*- coding: utf-8 -*-
"""
Regression tests for the dedup + decisions-log + push pipeline (dedup_check.py,
decisions_log.py, push_candidates.py). Pure, fast, zero LLM / network cost.

Run: python -m pytest tests/test_dedup_pipeline.py -v
"""

import json
import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dedup_check import is_recently_decided_duplicate, titles_match, companies_match, filter_new_jobs
from decisions_log import sync_from_dashboard_dump, record_pushed, load_decisions
from push_candidates import compute_doc_id

HERE = os.path.dirname(os.path.abspath(__file__))
PIPELINE_DIR = os.path.dirname(HERE)
DAY = 24 * 60 * 60 * 1000


# ── dedup_check.py ────────────────────────────────────────────────────────────

def test_title_exact_and_fuzzy_match():
    assert titles_match("Product Manager", "product manager")
    assert titles_match("Senior Product Manager - Supply Chain Domain", "Senior Product Manager - Supply Chain")
    assert not titles_match("Product Manager", "Software Engineer")


def test_company_fuzzy_match_handles_messy_scraped_fields():
    # Real case: a bad title/company split once produced "AI - Chennai - Kumaran Systems"
    # for what should just be "Kumaran Systems". Must still match.
    assert companies_match("AI - Chennai - Kumaran Systems", "Kumaran Systems Pvt Ltd")
    assert companies_match("FirstClub", "Firstclub")
    assert not companies_match("Kumaran Systems", "Cisco")


def test_decided_duplicate_blocked_within_window():
    now = int(time.time() * 1000)
    existing = [{"title": "Senior Product Manager", "company": "Acme", "user_action": "skip", "user_action_ts": now - 2 * DAY}]
    assert is_recently_decided_duplicate("Senior Product Manager", "Acme", existing) is True


def test_decided_duplicate_allowed_after_window():
    now = int(time.time() * 1000)
    existing = [{"title": "Data Analyst", "company": "OldCo", "user_action": "apply", "user_action_ts": now - 40 * DAY}]
    assert is_recently_decided_duplicate("Data Analyst", "OldCo", existing, window_days=30) is False


def test_never_decided_job_is_not_a_duplicate():
    existing = [{"title": "Product Manager", "company": "Random Co", "user_action": None, "user_action_ts": None}]
    assert is_recently_decided_duplicate("Product Manager", "Random Co", existing) is False


def test_same_company_different_title_is_not_a_duplicate():
    now = int(time.time() * 1000)
    existing = [{"title": "Totally Different Role", "company": "FirstClub", "user_action": "apply", "user_action_ts": now - 1 * DAY}]
    assert is_recently_decided_duplicate("Senior Product Manager - Supply Chain", "FirstClub", existing) is False


def test_filter_new_jobs_splits_keep_and_drop():
    now = int(time.time() * 1000)
    existing = [{"title": "Role A", "company": "CoA", "user_action": "apply", "user_action_ts": now - DAY}]
    candidates = [
        {"title": "Role A", "company": "CoA"},
        {"title": "Role B", "company": "CoB"},
    ]
    keep, dropped = filter_new_jobs(candidates, existing)
    assert [c["title"] for c in keep] == ["Role B"]
    assert [c["title"] for c in dropped] == ["Role A"]
    assert "_dedup_reason" in dropped[0]


# ── decisions_log.py (round-trips against a throwaway file, never the real log) ──

def test_decisions_log_sync_then_load_roundtrip(tmp_path):
    log_path = str(tmp_path / "test_log.xlsx")
    now_ms = int(time.time() * 1000)
    dump = [
        {"id": "job_aaa", "title": "Product Manager", "company": "Acme", "platform": "LinkedIn",
         "apply_url": "https://example.com/a", "user_action": "apply", "user_action_ts": now_ms, "date": "2026-11-26"},
        {"id": "job_bbb", "title": "Data Analyst", "company": "Beta", "platform": "Naukri",
         "apply_url": "https://example.com/b", "user_action": None, "user_action_ts": None, "date": "2026-11-26"},
    ]
    updated = sync_from_dashboard_dump(dump, log_path)
    assert updated == 1  # only job_aaa has a real decision
    assert os.path.exists(log_path)

    decisions = load_decisions(log_path)
    by_id = {d["doc_id"]: d for d in decisions}
    assert by_id["job_aaa"]["user_action"] == "apply"
    assert by_id["job_bbb"]["user_action"] is None


def test_decisions_log_never_loses_rows_on_resync(tmp_path):
    log_path = str(tmp_path / "test_log.xlsx")
    now_ms = int(time.time() * 1000)
    sync_from_dashboard_dump([{"id": "job_x", "title": "T", "company": "C", "platform": "P",
                                "apply_url": "u", "user_action": "skip", "user_action_ts": now_ms}], log_path)
    # Simulate the dashboard going away entirely (e.g. "Clear today's list") — syncing an
    # EMPTY dump must not erase what's already logged.
    sync_from_dashboard_dump([], log_path)
    decisions = load_decisions(log_path)
    assert len(decisions) == 1
    assert decisions[0]["doc_id"] == "job_x"


def test_record_pushed_adds_blank_status_rows(tmp_path):
    log_path = str(tmp_path / "test_log.xlsx")
    pushed = [{"doc_id": "job_new", "title": "New Role", "company": "NewCo", "platform": "LinkedIn", "apply_url": "https://example.com/new"}]
    added = record_pushed(pushed, log_path)
    assert added == 1
    decisions = load_decisions(log_path)
    assert decisions[0]["doc_id"] == "job_new"
    assert decisions[0]["user_action"] is None


# ── push_candidates.py end-to-end (subprocess, exercises the real CLI) ────────

def test_push_candidates_cli_end_to_end(tmp_path):
    log_path = str(tmp_path / "log.xlsx")
    now_ms = int(time.time() * 1000)
    sync_from_dashboard_dump([
        {"id": "job_decided", "title": "Decided Role", "company": "DecidedCo", "platform": "LinkedIn",
         "apply_url": "https://example.com/decided", "user_action": "apply", "user_action_ts": now_ms},
    ], log_path)
    record_pushed([
        {"doc_id": compute_doc_id("https://example.com/pushed"), "title": "Already Pushed Role", "company": "PushedCo",
         "platform": "Naukri", "apply_url": "https://example.com/pushed"},
    ], log_path)

    candidates = [
        {"title": "Decided Role", "company": "DecidedCo Pvt Ltd", "location": "Bengaluru",
         "apply_url": "https://example.com/decided-REPOSTED", "platform": "LinkedIn", "date": "2026-11-26", "jd_text": "", "note": None},
        {"title": "Already Pushed Role", "company": "PushedCo", "location": "Bengaluru",
         "apply_url": "https://example.com/pushed", "platform": "Naukri", "date": "2026-11-26", "jd_text": "", "note": None},
        {"title": "Brand New Role", "company": "FreshCo", "location": "Bengaluru",
         "apply_url": "https://example.com/fresh", "platform": "LinkedIn", "date": "2026-11-26", "jd_text": "", "note": None},
    ]
    candidates_path = str(tmp_path / "candidates.json")
    with open(candidates_path, "w", encoding="utf-8") as f:
        json.dump(candidates, f)
    out_path = str(tmp_path / "result.json")

    result = subprocess.run(
        [sys.executable, os.path.join(PIPELINE_DIR, "push_candidates.py"),
         "--candidates", candidates_path, "--decisions-log", log_path, "--out", out_path],
        capture_output=True, text=True, cwd=PIPELINE_DIR,
    )
    assert result.returncode == 0, result.stderr

    with open(out_path, encoding="utf-8") as f:
        output = json.load(f)
    pushed_titles = {c["title"] for c in output["to_push"]}
    dropped_titles = {c["title"] for c in output["dropped"]}
    assert pushed_titles == {"Brand New Role"}
    assert dropped_titles == {"Decided Role", "Already Pushed Role"}
