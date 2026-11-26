# -*- coding: utf-8 -*-
"""
Regression tests for job_scraper.py's pure helper functions — freshness filtering,
garbage-page detection, and the dashboard-override file consumption. No network calls,
no SerpApi credits spent. Importing job_scraper.py itself does real config.json
loading (needed for its module-level constants) but makes no outbound calls.

Run: python -m pytest tests/test_scraper_helpers.py -v
"""

import json
import os
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import job_scraper


# ── _naukri_posted_date / filter_naukri_freshness ─────────────────────────────

def test_naukri_posted_date_parses_real_url_format():
    # Real URL pushed to the dashboard today: trailing "061126000164" -> 06-11-26
    url = "https://www.naukri.com/job-listings-senior-product-manager-supply-chain-firstclub-bengaluru-5-to-8-years-061126000164"
    d = job_scraper._naukri_posted_date(url)
    assert d == date(2026, 11, 6)


def test_naukri_posted_date_none_for_unparseable_listicle_url():
    url = "https://www.naukri.com/5777-senior-product-manager-job-vacancies-in-bengaluru"
    assert job_scraper._naukri_posted_date(url) is None


def test_filter_naukri_freshness_drops_stale_and_unparseable():
    today = date.today()
    fresh_dd = (today - timedelta(days=5)).strftime("%d%m%y")
    stale_dd = (today - timedelta(days=60)).strftime("%d%m%y")
    jobs = [
        {"source_platform": "naukri", "apply_url": f"https://www.naukri.com/job-foo-{fresh_dd}123456"},
        {"source_platform": "naukri", "apply_url": f"https://www.naukri.com/job-bar-{stale_dd}123456"},
        {"source_platform": "naukri", "apply_url": "https://www.naukri.com/some-listicle-page"},
        {"source_platform": "linkedin", "apply_url": "https://linkedin.com/jobs/view/123"},  # not naukri — always kept
    ]
    kept = job_scraper.filter_naukri_freshness(jobs, max_age_days=21)
    kept_urls = {j["apply_url"] for j in kept}
    assert f"https://www.naukri.com/job-foo-{fresh_dd}123456" in kept_urls
    assert f"https://www.naukri.com/job-bar-{stale_dd}123456" not in kept_urls
    assert "https://www.naukri.com/some-listicle-page" not in kept_urls
    assert "https://linkedin.com/jobs/view/123" in kept_urls


# ── _is_linkedin_search_page ──────────────────────────────────────────────────

def test_linkedin_search_page_detector_catches_generic_results_page():
    garbage = "Where are the filters? You're now using AI-powered job search. 10,000+ Product Manager Jobs in Seattle"
    assert job_scraper._is_linkedin_search_page(garbage) is True


def test_linkedin_search_page_detector_does_not_flag_a_real_jd():
    # A real single-job JD page should never trip this, even with an unrelated sidebar
    # widget present ("Get notified about new jobs...") — that marker was tried and
    # dropped in an earlier version specifically because of this false-positive.
    real_jd = (
        "About the Job We are seeking a Technical Program Manager to lead complex projects. "
        "Responsibilities include identifying opportunities to automate manual processes. "
        "Get notified about new jobs in Bengaluru, Karnataka, India."
    )
    assert job_scraper._is_linkedin_search_page(real_jd) is False


# ── check_dashboard_override (monkeypatched to a throwaway file — never the real one) ──

def test_dashboard_override_custom_query_is_consumed_once(tmp_path, monkeypatch):
    override_path = str(tmp_path / "dashboard_override.json")
    monkeypatch.setattr(job_scraper, "OVERRIDE_FILE", override_path)
    with open(override_path, "w", encoding="utf-8") as f:
        json.dump({"custom_query": "ai procurement product manager bengaluru"}, f)

    result = job_scraper.check_dashboard_override()
    assert result is not None
    assert result["query"] == "ai procurement product manager bengaluru"
    assert not os.path.exists(override_path)  # consumed — must not fire twice


def test_dashboard_override_named_category_maps_to_config_entry(tmp_path, monkeypatch):
    override_path = str(tmp_path / "dashboard_override.json")
    monkeypatch.setattr(job_scraper, "OVERRIDE_FILE", override_path)
    real_name = job_scraper.CATEGORIES[0]["name"]
    with open(override_path, "w", encoding="utf-8") as f:
        json.dump({"category_name": real_name}, f)

    # save_category_state writes pipeline_state.json for real — isolate that too
    monkeypatch.setattr(job_scraper, "STATE_FILE", str(tmp_path / "pipeline_state.json"))
    result = job_scraper.check_dashboard_override()
    assert result == job_scraper.CATEGORIES[0]


def test_dashboard_override_missing_file_returns_none(tmp_path, monkeypatch):
    monkeypatch.setattr(job_scraper, "OVERRIDE_FILE", str(tmp_path / "does_not_exist.json"))
    assert job_scraper.check_dashboard_override() is None


def test_dashboard_override_unknown_category_name_is_ignored(tmp_path, monkeypatch):
    override_path = str(tmp_path / "dashboard_override.json")
    monkeypatch.setattr(job_scraper, "OVERRIDE_FILE", override_path)
    with open(override_path, "w", encoding="utf-8") as f:
        json.dump({"category_name": "Totally Made Up Category"}, f)
    assert job_scraper.check_dashboard_override() is None
