# -*- coding: utf-8 -*-
"""
dedup_check.py — title+company dedup for newly-scraped jobs, before pushing to the dashboard.

No LLM, no vector DB — just normalized string matching, per explicit user request (cost concern).

Rule (user's own words, 2026-11-26):
  A job is a duplicate ONLY if the same company+title was already shown on the dashboard AND the
  user acted on it (applied or skipped) within the last `window_days` (default 30). If they never
  acted on it, or they acted on it more than `window_days` ago, it's allowed to reappear — a listing
  can be genuinely reopened for a new hiring round a month later.

This module has no side effects — it's a pure filter. The caller is responsible for fetching
`existing_jobs` from the dashboard's `jobs` collection (ArtifactData query/list) before calling,
and for only pushing candidates that pass `is_new(...)`.
"""

import re
import time

_COMPANY_STOPWORDS = r"\b(pvt|ltd|limited|inc|llc|private|technologies|technology|solutions|corp|corporation|india)\b"


def normalize_company(c):
    c = (c or "").lower()
    c = re.sub(r"[^a-z0-9 ]", " ", c)
    c = re.sub(_COMPANY_STOPWORDS, "", c)
    c = re.sub(r"\s+", " ", c).strip()
    return c


def normalize_title(t):
    t = (t or "").lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _word_overlap_match(n1, n2, threshold=0.7):
    if not n1 or not n2:
        return False
    if n1 == n2:
        return True
    w1, w2 = set(n1.split()), set(n2.split())
    if not w1 or not w2:
        return False
    overlap = len(w1 & w2) / min(len(w1), len(w2))
    return overlap >= threshold


def titles_match(t1, t2):
    """Exact match after normalizing, or >=70% word overlap (catches platform-specific title variants
    like 'Senior Product Manager - Supply Chain Domain' vs 'Senior Product Manager')."""
    return _word_overlap_match(normalize_title(t1), normalize_title(t2), threshold=0.7)


def companies_match(c1, c2):
    """Exact match after normalizing, or >=70% word overlap. Real scraped company fields are often
    messy (e.g. a bad title/company split leaves location text glued on, like
    'AI - Chennai - Kumaran Systems' for what should just be 'Kumaran Systems') — a strict equality
    check misses these, so this uses the same word-overlap tolerance as titles_match."""
    return _word_overlap_match(normalize_company(c1), normalize_company(c2), threshold=0.7)


def is_recently_decided_duplicate(candidate_title, candidate_company, existing_jobs, window_days=30, now_ms=None):
    """
    existing_jobs: iterable of dicts, each with at least 'title', 'company', 'user_action',
                   'user_action_ts' (epoch milliseconds) — i.e. docs from the dashboard's
                   `jobs` collection.
    Returns True iff a title+company match exists whose user_action is 'apply' or 'skip' AND
    whose user_action_ts falls within the last `window_days` days of `now_ms` (defaults to now).
    """
    if now_ms is None:
        now_ms = int(time.time() * 1000)
    window_ms = window_days * 24 * 60 * 60 * 1000

    for job in existing_jobs:
        if job.get("user_action") not in ("apply", "skip"):
            continue
        ts = job.get("user_action_ts")
        if not ts or (now_ms - ts) > window_ms:
            continue
        if not companies_match(candidate_company, job.get("company")):
            continue
        if titles_match(candidate_title, job.get("title", "")):
            return True
    return False


def filter_new_jobs(candidates, existing_jobs, window_days=30, now_ms=None):
    """
    candidates: list of dicts each with 'title' and 'company' (plus whatever other fields
                the caller wants to push through unchanged).
    Returns (keep, dropped) — keep: candidates that pass the dedup check and should be pushed;
    dropped: candidates filtered out, each with a '_dedup_reason' key added for logging.
    """
    keep, dropped = [], []
    for c in candidates:
        if is_recently_decided_duplicate(c.get("title", ""), c.get("company", ""), existing_jobs, window_days, now_ms):
            c = dict(c)
            c["_dedup_reason"] = f"already {('decided on')} within {window_days} days — same company+title"
            dropped.append(c)
        else:
            keep.append(c)
    return keep, dropped
