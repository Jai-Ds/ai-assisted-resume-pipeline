# -*- coding: utf-8 -*-
"""
job_scraper.py — v2 (multi-platform, category rotation)

6-Stage Pipeline:
  1. Fetch   — SerpApi: LinkedIn, Hirist, Indeed, Naukri (google engine) + Google Jobs engine
  2. Filter  — Google Jobs via-field dedup + low-quality aggregator drop (jobrapido/bebee/
               simplyhired), then Naukri URL-date freshness filter (see filter_naukri_freshness)
  3. Blacklist — Drop noise titles (intern, director, VP, etc.)
  4. Dedup   — Hard (job_id) + Soft (cross-platform same role → `claude -p` confirms)
  5. JD Fetch — WebFetch for LinkedIn, Playwright for Hirist (SPA, JS-rendered)/Indeed/Google
               Jobs. Hirist listings are also freshness-checked via their own "Posted on" date;
               an absent date means the listing has expired (site falls back to homepage content).
  6. Output  — staged_jobs/ JSON files + Daily_Jobs_Log.xlsx

Category rotation: cycles through config.json's `categories` list, one per day, then repeats.
Budget: ~6 SerpApi credits/run (1 per platform, 2 for Google Jobs) — size your plan accordingly.
Invocation: scheduled via Windows Task Scheduler / cron, or run manually.

All user-specific values (API key, search categories/queries, title blacklist, paths) live in
config.json — see config.example.json and the README for setup.
"""

import os
import sys
import json
import re
import time
from datetime import datetime, date
from urllib.parse import urlsplit

# Force UTF-8 output — Windows console defaults to cp1252 which can't handle arrows/ticks
sys.stdout.reconfigure(encoding="utf-8")

from serpapi import GoogleSearch
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from config import SERPAPI_KEY, PATHS, CATEGORIES, TITLE_BLACKLIST, PROFILE


# ── Config ──────────────────────────────────────────────────────────────────
# All of the above — the SerpApi key, every path, the search categories/queries,
# and the title blacklist — are user-specific and live in config.json (copy
# config.example.json to get started). Nothing in this file is hardcoded to
# any one person's job search; repoint config.json and this pipeline searches
# for whatever role you're after.

PIPELINE_DIR   = PATHS["pipeline_dir"]
STAGED_DIR     = os.path.join(PIPELINE_DIR, "staged_jobs")
PROCESSED_FILE = os.path.join(PIPELINE_DIR, "processed_jobs.json")
DAILY_LOG_FILE = os.path.join(PIPELINE_DIR, "Daily_Jobs_Log.xlsx")
STATE_FILE     = os.path.join(PIPELINE_DIR, "pipeline_state.json")
RUN_LOG_FILE   = os.path.join(PIPELINE_DIR, "scraper_run_log.xlsx")


def get_today_category():
    """
    Compute today's category index based on pipeline_state.json.
    Returns (idx, is_same_day_rerun) — does NOT save state yet.
    Call save_category_state(idx) after user confirms.
    """
    today_str = date.today().isoformat()

    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
    else:
        state = {"last_run_date": None, "category_index": 0}

    last_run = state.get("last_run_date")
    idx      = state.get("category_index", 0)

    if last_run == today_str:
        same_day = True   # Same day re-run — keep same category
    elif last_run is None:
        idx      = 0      # First ever run → Day 1 (Supply Chain)
        same_day = False
    else:
        idx      = (idx + 1) % len(CATEGORIES)  # Advance to next category
        same_day = False

    return idx, same_day


def save_category_state(idx):
    """Persist the chosen category index and today's date to pipeline_state.json."""
    today_str = date.today().isoformat()
    state = {"last_run_date": today_str, "category_index": idx}
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


OVERRIDE_FILE = os.path.join(PIPELINE_DIR, "dashboard_override.json")


def check_dashboard_override():
    """
    Look for a one-off override dropped locally (by a live Claude Code session that
    relayed a dashboard button click — this script itself never talks to the dashboard).
    File shape: {"category_name": "AI Supply Chain"} to jump to a named config.json
    category, or {"custom_query": "..."} to search exactly that text once. Consumed
    (deleted) immediately so it never fires twice. Returns a category dict or None.
    """
    if not os.path.exists(OVERRIDE_FILE):
        return None
    try:
        with open(OVERRIDE_FILE, encoding="utf-8") as f:
            override = json.load(f)
    except Exception as e:
        print(f"  WARNING: dashboard_override.json unreadable ({e}) — ignoring it.")
        os.remove(OVERRIDE_FILE)
        return None
    os.remove(OVERRIDE_FILE)  # one-shot — consume before acting, never re-fire on crash/retry

    custom_query = (override.get("custom_query") or "").strip()
    if custom_query:
        print(f"\n  Dashboard override: custom search queued — \"{custom_query}\"")
        print("  (one-off — today's run does not count toward or disturb the normal rotation)")
        return {"name": "Custom (dashboard)", "query": custom_query}

    cat_name = (override.get("category_name") or "").strip()
    if cat_name:
        for i, cat in enumerate(CATEGORIES):
            if cat["name"] == cat_name:
                print(f"\n  Dashboard override: jumping to \"{cat_name}\"")
                save_category_state(i)  # counts as today's rotation turn, same as a manual menu pick
                return cat
        print(f"  WARNING: dashboard_override.json named category \"{cat_name}\" — not found in config.json's categories, ignoring.")
    return None


def confirm_category(computed_idx, same_day):
    """
    Show the queued category and full list. Let user confirm with Enter
    or override by typing 1-4. Returns the final chosen category dict.
    """
    print(f"\n{'='*62}")
    print(f"  TODAY'S SEARCH")
    print(f"{'='*62}")
    if same_day:
        print(f"  (Same-day re-run — category unchanged)")
    print()
    for i, cat in enumerate(CATEGORIES):
        marker = "  <-- queued" if i == computed_idx else ""
        print(f"    [{i+1}] {cat['name']}{marker}")
    print()

    while True:
        try:
            raw = input(f"  Press ENTER to run [{computed_idx+1}] {CATEGORIES[computed_idx]['name']}, "
                        f"or type 1-{len(CATEGORIES)} to override: ").strip()
        except EOFError:
            # No one at the keyboard (e.g. Task Scheduler with stdin redirected from NUL) —
            # behave exactly like pressing ENTER instead of crashing before any scraping happens.
            print("  (no input available — running unattended, using queued category)")
            chosen_idx = computed_idx
            break
        if raw == "":
            chosen_idx = computed_idx
            break
        if raw.isdigit() and 1 <= int(raw) <= len(CATEGORIES):
            chosen_idx = int(raw) - 1
            if chosen_idx != computed_idx:
                print(f"  Overriding to [{chosen_idx+1}] {CATEGORIES[chosen_idx]['name']}")
            break
        print(f"  Invalid — enter a number between 1 and {len(CATEGORIES)}, or press ENTER.")

    save_category_state(chosen_idx)
    return CATEGORIES[chosen_idx]


# ── Constants ────────────────────────────────────────────────────────────────
# TITLE_BLACKLIST comes from config.json (see import above) — tune it for your
# own search there, not here.

# Google Jobs `via` values to drop: known platforms already fetched directly
# (dedup), PLUS low-quality reposting aggregators (jobrapido, bebee, simplyhired)
# that republish old/closed listings from elsewhere with a fabricated "just
# posted" timestamp — a known source of stale/dead links slipping past the
# date_posted:week filter.
EXCLUDE_VIA = [
    "linkedin", "naukri", "indeed", "hirist", "glassdoor", "shine",
    "jobrapido", "bebee", "simplyhired",
]

# Listings older than this are dropped even if a platform's own freshness
# filter (tbs=qdr / date_posted) let them through — see filter_naukri_freshness()
# and the Hirist branch of enrich_jd().
MAX_LISTING_AGE_DAYS = 21

HEADERS_WEB = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


# ── Helpers ──────────────────────────────────────────────────────────────────

def sanitize_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "", text or "").strip().replace(" ", "_")[:50]


def load_processed():
    if os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_processed(ids):
    with open(PROCESSED_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(ids), f, indent=2)


def strip_html(html):
    """Strip HTML tags and collapse whitespace."""
    text = re.sub(r'<script[^>]*>.*?</script>', ' ', html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<style[^>]*>.*?</style>',  ' ', text,  flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'&nbsp;', ' ', text)
    text = re.sub(r'&amp;',  '&', text)
    text = re.sub(r'&lt;',   '<', text)
    text = re.sub(r'&gt;',   '>', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def clean_linkedin_jd(text):
    """
    LinkedIn logged-out pages have login-wall noise at BOTH the top (before JD)
    and the bottom (after JD). We want to strip only the top preamble.

    Two top-preamble types observed:
      Active job  — has 'Sign in to tailor your resume' etc. (~2200 chars of top noise)
      Expired job — simpler, ends with 'Report this job' (~1500 chars of top noise)

    Strategy: search only the first 40% of the page for known noise markers.
    This avoids accidentally picking up the identical sign-in prompts at the bottom.
    """
    noise_patterns = [
        r'Report this job',                        # present on all page types
        r'Sign in to tailor your resume',          # active jobs only
        r'Sign in to evaluate your skills',        # active jobs only
        r'Sign in to access AI-powered',           # active jobs only
        r'By clicking Continue to join or sign in',
        r'New to LinkedIn\? Join now',
    ]
    search_zone = text[:int(len(text) * 0.40)]
    noise_end = 0
    for pattern in noise_patterns:
        for m in re.finditer(pattern, search_zone, re.IGNORECASE):
            if m.end() > noise_end:
                noise_end = m.end()

    if noise_end > 100:
        cleaned = text[noise_end:].strip()
        if len(cleaned) > 500:
            return cleaned

    return text


# ── Stage 1: Fetch ───────────────────────────────────────────────────────────

def _extract_job_id(url, platform):
    """Extract a stable job ID from a job URL."""
    if platform == "linkedin":
        m = re.search(r'/jobs/view/(\d+)', url)
        return f"li_{m.group(1)}" if m else url
    if platform == "hirist":
        m = re.search(r'/j/[^/]+-(\d+)', url)
        return f"hi_{m.group(1)}" if m else url
    if platform == "indeed":
        m = re.search(r'jk=([a-z0-9]+)', url, re.IGNORECASE)
        return f"in_{m.group(1)}" if m else url
    return url


def _parse_title_company(raw, platform):
    """
    Best-effort parse of role title + company from a Google search result title.

    Platform behaviour (observed):
      LinkedIn : "Role at Company - Location"  or  "Company hiring Role"
      Hirist   : "Company - Role Title"         (dash order is REVERSED vs LinkedIn)
      Indeed   : "Role Title - City, State"     (no company in title — extracted from JD later)
    """
    for suffix in [" | LinkedIn", "| LinkedIn Jobs", " | Hirist", "| Hirist Jobs",
                   " | Indeed", "| Indeed Jobs"]:
        raw = raw.replace(suffix, "")
    raw = raw.strip()

    # "Role at Company" — works for LinkedIn
    m = re.match(r'^(.+?)\s+at\s+(.+)$', raw, re.IGNORECASE)
    if m:
        return m.group(1).strip(), m.group(2).strip()

    # "Company hiring Role" — works for LinkedIn/Hirist
    m = re.match(r'^(.+?)\s+(?:is\s+)?hiring\s+(?:a\s+|an\s+)?(.+)$', raw, re.IGNORECASE)
    if m:
        return m.group(2).strip(), m.group(1).strip()

    # Dash split — direction differs by platform
    parts = raw.split(" - ", 1)
    if len(parts) == 2:
        if platform == "hirist":
            # Hirist: "Company - Role Title"
            return parts[1].strip(), parts[0].strip()
        elif platform == "indeed":
            # Indeed: "Role Title - City, State" — second part is location, not company
            return parts[0].strip(), ""
        else:
            # LinkedIn fallback: "Role - Company/Location"
            return parts[0].strip(), parts[1].strip()

    return raw, ""



def fetch_google_site(label, site_restrict, query, platform_tag, pages=1):
    """
    1 SerpApi credit per page. Google engine with site: restriction.
    pages=2 pulls a second page (results 11-20) via the "start" offset —
    unlike the google_jobs engine, the classic "google" engine's pagination
    is via "start", not next_page_token (that rule is specific to google_jobs).
    Returns list of normalized job dicts.
    """
    # Google silently drops the site: restriction (returning ordinary organic
    # results for the leftover query text) when too few/no pages actually
    # match that site. SerpApi passes those through unfiltered, so without
    # this check, off-domain junk (Wikipedia/Reddit/gov-site hits for stray
    # query words like "bengaluru") gets treated as real job leads. Require
    # the result's own domain to actually match site_restrict's domain.
    site_domain = site_restrict.split("/")[0].lower()

    jobs       = []
    off_domain = 0

    for page in range(pages):
        params = {
            "engine":  "google",
            "q":       f"{query} site:{site_restrict}",
            # NOTE: tbs=qdr is Google's crawl/index-freshness signal, not the job's
            # actual posting date — pages Google merely re-crawled recently (e.g.
            # Naukri's evergreen listicle/category pages) can pass this even when
            # the individual job inside is months old. Kept tight (2d, was 4d) to
            # reduce that window; real freshness enforcement happens downstream
            # (filter_naukri_freshness, the Hirist posted-date check in enrich_jd).
            "tbs":     "qdr:2d",
            "gl":      "in",
            "hl":      "en",
            "num":     10,
            "start":   page * 10,
            "api_key": SERPAPI_KEY,
        }
        search  = GoogleSearch(params)
        results = search.get_dict()

        if "error" in results:
            print(f"  [{label}] Page {page+1} SerpApi error: {results['error']}")
            break

        organic = results.get("organic_results", [])
        print(f"  [{label}] Page {page+1}: {len(organic)} result(s)")

        for r in organic:
            url       = r.get("link", "")
            title_raw = r.get("title", "")
            snippet   = r.get("snippet", "")

            netloc = urlsplit(url).netloc.lower()
            if site_domain not in netloc:
                off_domain += 1
                continue

            title, company = _parse_title_company(title_raw, platform_tag)
            job_id = _extract_job_id(url, platform_tag)

            jobs.append({
                "job_id":          job_id or url,
                "title":           title,
                "company":         company,
                "location":        "",
                "description":     snippet,
                "apply_url":       url,
                "source_link":     "",
                "via":             "",
                "source_platform": platform_tag,
                "fetched_at":      datetime.now().strftime("%Y-%m-%d %H:%M"),
                "has_jd":          False,
            })

        if not organic:
            break  # no more results — don't burn a credit on an empty next page

    if off_domain:
        print(f"  [{label}] dropped {off_domain} off-domain result(s) — site: restriction didn't hold")

    return jobs


def fetch_google_jobs(query):
    """
    2 SerpApi credits (2 pages). google_jobs engine.
    Returns list of normalized job dicts.
    """
    jobs       = []
    next_token = None

    for page in range(2):
        params = {
            "engine":  "google_jobs",
            "q":       query,
            "gl":      "in",
            "hl":      "en",
            "chips":   "date_posted:week",
            "api_key": SERPAPI_KEY,
        }
        if next_token:
            params["next_page_token"] = next_token

        search  = GoogleSearch(params)
        results = search.get_dict()

        if "error" in results:
            print(f"  [Google Jobs] Page {page+1} error: {results['error']}")
            break

        raw = results.get("jobs_results", [])
        print(f"  [Google Jobs] Page {page+1}: {len(raw)} result(s)")

        for j in raw:
            apply_opts  = j.get("apply_options", [])
            apply_url   = apply_opts[0]["link"] if apply_opts else j.get("source_link", "")
            source_link = j.get("source_link", "")
            desc        = (j.get("description") or "").strip()

            jobs.append({
                "job_id":          j.get("job_id", ""),
                "title":           j.get("title", ""),
                "company":         j.get("company_name", ""),
                "location":        j.get("location", ""),
                "description":     desc,
                "apply_url":       apply_url,
                "source_link":     source_link,
                "via":             j.get("via", ""),
                "source_platform": "google_jobs",
                "fetched_at":      datetime.now().strftime("%Y-%m-%d %H:%M"),
                "has_jd":          len(desc) > 200,
            })

        next_token = results.get("serpapi_pagination", {}).get("next_page_token")
        if not next_token:
            break

    return jobs


# ── Stage 2: Google Jobs via filter ─────────────────────────────────────────

def filter_google_jobs_via(jobs):
    """
    Remove Google Jobs entries that came via known platforms already fetched directly.
    -site: operator does NOT work on google_jobs engine — this Python filter is the fix.
    """
    before   = len(jobs)
    filtered = [
        j for j in jobs
        if j.get("source_platform") != "google_jobs"
        or not any(x in j.get("via", "").lower() for x in EXCLUDE_VIA)
    ]
    removed = before - len(filtered)
    print(f"  [via filter] {before} → {len(filtered)} ({removed} platform duplicates removed)")
    return filtered


def _naukri_posted_date(url):
    """
    Naukri job URLs encode the posting date in the trailing numeric ID as
    DDMMYY + a 6-digit serial (e.g. ...-020925501157 → 02-09-25). JD fetch is
    permanently blocked (Akamai), but the date is readable from the URL alone
    without hitting the site, so listings can still be freshness-filtered for
    manual follow-up. Returns a date, or None if the pattern doesn't match.
    """
    m = re.search(r'-(\d{6})\d{6}$', urlsplit(url).path)
    if not m:
        return None
    dd, mm, yy = int(m.group(1)[0:2]), int(m.group(1)[2:4]), int(m.group(1)[4:6])
    try:
        return date(2000 + yy, mm, dd)
    except ValueError:
        return None


def filter_naukri_freshness(jobs, max_age_days=MAX_LISTING_AGE_DAYS):
    """
    Drop Naukri listings whose URL-encoded posting date is older than
    max_age_days, or unparseable listicle/category pages (e.g. "5777 Senior
    Product Manager Job Vacancies In Bengaluru") that aren't single postings
    at all. Naukri stays in discovery (per user request — JD can't be
    fetched, but a recent match is still worth opening manually), this just
    keeps what surfaces actually worth clicking into.
    """
    before  = len(jobs)
    kept    = []
    dropped = 0
    for j in jobs:
        if j.get("source_platform") != "naukri":
            kept.append(j)
            continue
        posted = _naukri_posted_date(j.get("apply_url", ""))
        if posted is None:
            dropped += 1  # unparseable → almost always a category/listicle page, not a real posting
            continue
        age = (date.today() - posted).days
        if -7 <= age <= max_age_days:
            kept.append(j)
        else:
            dropped += 1
    print(f"  [naukri freshness] {before} → {len(kept)} ({dropped} stale/unparseable Naukri listing(s) dropped)")
    return kept


# ── Stage 3: Title blacklist ─────────────────────────────────────────────────

def apply_title_blacklist(jobs):
    before = len(jobs)
    kept   = [j for j in jobs
              if not any(b in (j.get("title") or "").lower() for b in TITLE_BLACKLIST)]
    print(f"  [blacklist] {before} → {len(kept)} ({before - len(kept)} removed)")
    return kept


# ── Stage 4: Deduplication ───────────────────────────────────────────────────

def hard_dedup(jobs, processed_ids):
    """Remove jobs whose job_id was seen in a previous run."""
    new = [j for j in jobs if j.get("job_id") not in processed_ids]
    print(f"  [hard dedup] {len(jobs)} → {len(new)} new ({len(jobs) - len(new)} already seen)")
    return new


def soft_dedup(jobs):
    """
    Cross-platform soft dedup: if the same (company, title) appears from 2 different
    platforms today, call Claude to confirm if it's the same role.
    Only fires on suspicious pairs — not every job.
    """
    seen       = {}  # normalized (company, title) → first job index
    suspicious = []  # list of (job_a, job_b) index pairs

    for i, j in enumerate(jobs):
        key = (
            re.sub(r'\W+', '', (j.get("company") or "").lower()),
            re.sub(r'\W+', '', (j.get("title")   or "").lower()),
        )
        if not key[0] or not key[1]:
            continue
        if key in seen:
            other_idx = seen[key]
            if jobs[other_idx].get("source_platform") != j.get("source_platform"):
                suspicious.append((other_idx, i))
        else:
            seen[key] = i

    if not suspicious:
        print(f"  [soft dedup] No suspicious cross-platform pairs.")
        return jobs

    print(f"  [soft dedup] {len(suspicious)} suspicious pair(s) — checking with Claude...")

    to_remove = set()
    for idx_a, idx_b in suspicious:
        a, b = jobs[idx_a], jobs[idx_b]
        if _claude_same_job(a, b):
            # Keep whichever has the longer (richer) description
            drop = b if len(a.get("description", "")) >= len(b.get("description", "")) else a
            to_remove.add(id(drop))
            print(f"    Duplicate → kept {a['source_platform']}, "
                  f"dropped {drop['source_platform']}: {a['company']} / {a['title']}")
        else:
            print(f"    Different roles: [{a['source_platform']}] {a['title']} @ {a['company']} "
                  f"vs [{b['source_platform']}] {b['title']} @ {b['company']}")

    cleaned = [j for j in jobs if id(j) not in to_remove]
    print(f"  [soft dedup] {len(jobs)} → {len(cleaned)} ({len(to_remove)} removed)")
    return cleaned


def _claude_same_job(a, b):
    """
    Ask Claude via `claude -p` (Claude Code CLI) if two listings are the same role.
    Falls back to False (keep both) on any error.
    """
    import subprocess

    prompt = (
        f"Job A: {a.get('title')} at {a.get('company')} "
        f"({a.get('location')}) via {a.get('source_platform')}\n"
        f"Job B: {b.get('title')} at {b.get('company')} "
        f"({b.get('location')}) via {b.get('source_platform')}\n\n"
        "Are these the same job posting cross-listed on two platforms? "
        "Reply SAME or DIFFERENT only."
    )

    try:
        result = subprocess.run(
            ["cmd", "/c", "claude", "-p", prompt],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return "SAME" in result.stdout.upper()
    except Exception as e:
        print(f"    claude -p error: {e} — keeping both")
        return False


# ── Stage 5: JD Fetching ─────────────────────────────────────────────────────

def _is_linkedin_search_page(text):
    """
    LinkedIn sometimes serves a generic multi-job search-results page instead
    of the specific listing (e.g. the job was removed or geo-redirected) —
    a plain requests.get() follows the redirect and gets a 200, so the old
    code accepted it as a real JD. These markers are LinkedIn's own generic
    search-UI boilerplate and don't appear on an actual single-job JD page.
    """
    markers = [
        r'Where are the filters\?',
        r"You're now using AI-powered job search",
        r'\d[\d,]*\+\s+[\w\s/&-]{3,60}?\s+[Jj]obs? in\b',
    ]
    return any(re.search(m, text) for m in markers)


def fetch_jd_webfetch(url):
    """Fetch and clean JD text via requests. Used for LinkedIn and Hirist."""
    try:
        resp = requests.get(url, headers=HEADERS_WEB, timeout=15, allow_redirects=True)
        if resp.status_code == 200:
            text = strip_html(resp.text)
            return text[:6000] if len(text) > 150 else ""
        print(f"    HTTP {resp.status_code} for {url[:70]}")
        return ""
    except Exception as e:
        print(f"    WebFetch error: {e}")
        return ""


def fetch_jd_playwright(url, wait_s=2):
    """
    Fetch JD via headless Playwright.
    Used for Indeed (stealth), Google Jobs source_link, and Hirist (SPA —
    needs a longer wait_s than the others for its JS to hydrate).
    """
    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/122.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 800},
                locale="en-US",
                extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
            )
            page = ctx.new_page()

            # Block images/fonts/media for speed
            page.route(
                "**/*.{png,jpg,jpeg,gif,webp,svg,woff,woff2,ttf,otf,mp4,mp3}",
                lambda route: route.abort(),
            )

            page.goto(url, timeout=25000, wait_until="domcontentloaded")
            time.sleep(wait_s)

            text = page.inner_text("body")
            browser.close()

            text = re.sub(r'\s+', ' ', text).strip()
            return text[:6000] if len(text) > 150 else ""

    except Exception as e:
        print(f"    Playwright error: {e}")
        return ""


def _hirist_posted_date(text):
    """
    Live Hirist listings show 'Posted on: DD/MM/YYYY'. Expired listings render
    fine (no HTTP error) but silently fall back to generic homepage content
    with NO 'Posted on' marker at all — so its absence is used as an expiry
    signal, not just a parse failure.
    """
    m = re.search(r'Posted on:?\s*(\d{1,2})/(\d{1,2})/(\d{4})', text)
    if not m:
        return None
    dd, mm, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(yyyy, mm, dd)
    except ValueError:
        return None


def enrich_jd(jobs):
    """Stage 5: Add full JD text to each job based on its source platform."""
    print(f"\nStage 5: Fetching JDs for {len(jobs)} job(s)...")

    for j in jobs:
        platform = j.get("source_platform", "")

        # Skip if already has a substantial JD (e.g. from Google Jobs engine)
        if j.get("has_jd") and len(j.get("description", "")) > 300:
            print(f"  [skip] {j.get('company','?')} — already has JD "
                  f"({len(j['description'])} chars)")
            continue

        company = j.get("company") or j.get("title") or "?"
        title   = j.get("title", "?")
        print(f"  Fetching: {company} / {title} [{platform}]", end="", flush=True)

        if platform == "linkedin":
            jd = fetch_jd_webfetch(j.get("apply_url", ""))
            if jd:
                jd = clean_linkedin_jd(jd)
                if _is_linkedin_search_page(jd):
                    print(" [LinkedIn — redirected to generic search results, not a real JD]", end="")
                    jd = ""
        elif platform == "hirist":
            # Hirist now renders JDs client-side — plain requests only see a
            # "You need to enable JavaScript" shell, so this needs Playwright.
            jd     = fetch_jd_playwright(j.get("apply_url", ""), wait_s=5)
            posted = _hirist_posted_date(jd) if jd else None
            if not jd or posted is None:
                j["description"] = "No JD available — Hirist listing appears expired"
                j["has_jd"]      = False
                print(" [Hirist — listing appears expired or unreachable]")
                continue
            age = (date.today() - posted).days
            if age > MAX_LISTING_AGE_DAYS or age < -7:
                j["description"] = f"No JD available — Hirist listing too old (posted {posted.isoformat()})"
                j["has_jd"]      = False
                print(f" [Hirist — stale, posted {posted.isoformat()}]")
                continue
            j["description"] = jd
            j["has_jd"]      = True
            print(f" ✓ {len(jd)} chars (posted {posted.isoformat()})")
            continue
        elif platform == "indeed":
            # Indeed uses Cloudflare — Playwright blocked at scale.
            # Keep the job for discovery (URL + title), skip JD fetch.
            j["description"] = "No JD available — check Apply URL"
            j["has_jd"]      = False
            print(" [Indeed — Cloudflare blocks JD fetch, check URL manually]")
            continue
        elif platform == "naukri":
            # Naukri uses Akamai — try WebFetch (lighter than Playwright, sometimes slips through).
            # Fall back to URL-only if blocked.
            jd = fetch_jd_webfetch(j.get("apply_url", ""))
            if not jd:
                j["description"] = "No JD available — check Apply URL (Naukri/Akamai)"
                j["has_jd"]      = False
                print(" [Naukri — Akamai blocked, check URL manually]")
                continue
        elif platform == "google_jobs":
            # Prefer source_link (the origin site) over the apply link
            url = j.get("source_link") or j.get("apply_url", "")
            jd  = fetch_jd_playwright(url)
        else:
            jd = fetch_jd_webfetch(j.get("apply_url", ""))

        if jd:
            j["description"] = jd
            j["has_jd"]      = True
            print(f" ✓ {len(jd)} chars")
        else:
            j["description"] = "No JD available — check Apply URL"
            j["has_jd"]      = False
            print(" ✗ no JD — flagged for manual review")

    return jobs


# ── Stage 6: Output ──────────────────────────────────────────────────────────

def stage_jobs(jobs):
    """Write each job as a .json file in staged_jobs/."""
    os.makedirs(STAGED_DIR, exist_ok=True)
    today  = datetime.now().strftime("%Y-%m-%d")
    saved  = []

    for j in jobs:
        company  = sanitize_filename(j.get("company", "Unknown"))
        title    = sanitize_filename(j.get("title",   "PM"))
        filename = f"{today}_{company}_{title}.json"
        filepath = os.path.join(STAGED_DIR, filename)

        payload = {
            "job_id":          j.get("job_id", ""),
            "title":           j.get("title", ""),
            "company":         j.get("company", ""),
            "location":        j.get("location", ""),
            "description":     j.get("description", ""),
            "apply_url":       j.get("apply_url", ""),
            "source_platform": j.get("source_platform", ""),
            "fetched_at":      j.get("fetched_at", ""),
            "has_jd":          j.get("has_jd", False),
            "sheet_logged":    False,
        }

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)

        saved.append(filepath)
        print(f"  Staged: {filename}")

    return saved


# ── Staged job sheet-logging helpers ─────────────────────────────────────────

def get_unlogged_staged():
    """Return list of (fpath, job_dict) for staged jobs not yet written to sheet."""
    unlogged = []
    if not os.path.isdir(STAGED_DIR):
        return unlogged
    for fname in sorted(os.listdir(STAGED_DIR)):
        if not fname.endswith(".json"):
            continue
        fpath = os.path.join(STAGED_DIR, fname)
        try:
            with open(fpath, encoding="utf-8") as f:
                job = json.load(f)
            if not job.get("sheet_logged"):
                unlogged.append((fpath, job))
        except Exception:
            continue
    return unlogged


def mark_staged_logged(fpaths):
    """Set sheet_logged=True in the given staged JSON files."""
    for fpath in fpaths:
        try:
            with open(fpath, encoding="utf-8") as f:
                job = json.load(f)
            job["sheet_logged"] = True
            with open(fpath, "w", encoding="utf-8") as f:
                json.dump(job, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"  WARNING: Could not mark {os.path.basename(fpath)} as logged: {e}")


# ── Daily Log Excel ──────────────────────────────────────────────────────────

HEADERS    = ["Date Fetched", "Company", "Role Title", "Location",
              "Apply URL", "Full JD", "Status", "Date Applied", "Notes",
              "Base Resume", "Runner Status"]
COL_WIDTHS = [14, 22, 38, 18, 45, 90, 14, 14, 30, 14, 14]

DARK_BLUE  = "1F3864"
MID_BLUE   = "2E75B6"
LIGHT_BLUE = "DEEAF1"
ALT_ROW    = "F5F9FD"
WHITE      = "FFFFFF"


def _center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def _left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def _fill(c):
    return PatternFill("solid", fgColor=c)


def update_daily_log(new_jobs):
    """Append new rows to Daily_Jobs_Log.xlsx. Creates file with headers if missing."""

    lock = os.path.join(PIPELINE_DIR, "~$Daily_Jobs_Log.xlsx")
    if os.path.exists(lock):
        print("  WARNING: Daily_Jobs_Log.xlsx is open in Excel. Close it and re-run to log.")
        return

    if os.path.exists(DAILY_LOG_FILE):
        wb = openpyxl.load_workbook(DAILY_LOG_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Job Leads"

        # Title banner
        ws.merge_cells("A1:K1")
        ws["A1"]           = f"{PROFILE['name']} — Daily Job Leads"
        ws["A1"].font      = Font(bold=True, size=14, color=WHITE)
        ws["A1"].fill      = _fill(DARK_BLUE)
        ws["A1"].alignment = _center()
        ws.row_dimensions[1].height = 28

        # Column headers
        for col, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
            c            = ws.cell(row=2, column=col, value=h)
            c.font       = Font(bold=True, color=WHITE, size=10)
            c.fill       = _fill(MID_BLUE)
            c.alignment  = _center()
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.row_dimensions[2].height = 22
        ws.freeze_panes = "A3"

    # Status dropdown (col 7)
    dv = DataValidation(
        type="list",
        formula1='"New,Shortlisted,Applied,Interview,Rejected,Skipped"',
        allow_blank=True,
        showDropDown=False,
    )
    ws.add_data_validation(dv)

    # Runner Status dropdown (col 11)
    dv_runner = DataValidation(
        type="list",
        formula1='"Done,Failed,Skipped-NoJD"',
        allow_blank=True,
        showDropDown=False,
    )
    ws.add_data_validation(dv_runner)

    today = datetime.now().strftime("%Y-%m-%d")

    for j in new_jobs:
        apply_url = j.get("apply_url", "")
        full_jd   = (j.get("description") or "No JD available — check Apply URL").strip()

        ws.append([
            today,
            j.get("company", ""),
            j.get("title", ""),
            j.get("location", ""),
            apply_url,
            full_jd,
            "New",
            "",
            "",
            "",   # Base Resume — filled by job_runner.py
            "",   # Runner Status — filled by job_runner.py
        ])
        r = ws.max_row

        row_fill = _fill(LIGHT_BLUE) if r % 2 == 0 else _fill(ALT_ROW)

        for col in range(1, len(HEADERS) + 1):
            cell = ws.cell(r, col)
            cell.fill = row_fill
            cell.font = Font(size=9)
            if col in (1, 7, 8, 10, 11):  # Date Fetched, Status, Date Applied, Base Resume, Runner Status
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif col == 6:             # Full JD — top-aligned for readability
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # Apply URL as clickable hyperlink
        if apply_url:
            url_cell           = ws.cell(r, 5)
            url_cell.value     = apply_url
            url_cell.hyperlink = apply_url
            url_cell.font      = Font(size=8, color="0563C1", underline="single")
            url_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        dv.add(ws.cell(r, 7))
        dv_runner.add(ws.cell(r, 11))
        ws.row_dimensions[r].height = 120

    wb.save(DAILY_LOG_FILE)
    print(f"  Daily log saved → {DAILY_LOG_FILE}")


# ── Run Log ──────────────────────────────────────────────────────────────────

RUN_LOG_HEADERS = [
    "Date", "Time", "Category",
    "LinkedIn Fetched", "LinkedIn JD OK", "LinkedIn JD Fail",
    "Hirist Fetched",  "Hirist JD OK",  "Hirist JD Fail",
    "Indeed Fetched",  "Indeed JD OK",  "Indeed JD Fail",
    "GJobs Fetched",   "GJobs JD OK",   "GJobs JD Fail",
    "Naukri Fetched",  "Naukri JD OK",  "Naukri JD Fail",
    "Raw Total", "After Filters", "New (post-dedup)", "Staged",
    "Notes",
]
RUN_LOG_WIDTHS = [12, 8, 14, 14, 12, 12, 14, 12, 12, 14, 12, 12, 12, 10, 10, 12, 10, 10, 10, 14, 16, 8, 40]


def write_run_log(category, all_jobs_raw, all_jobs_filtered, all_jobs_new, staged_count, notes=""):
    """Append one row to scraper_run_log.xlsx with per-platform JD success/fail counts."""

    now = datetime.now()

    # Per-platform stats from the final staged jobs
    stats = {}
    for p in ("linkedin", "hirist", "indeed", "google_jobs", "naukri"):
        platform_jobs = [j for j in all_jobs_new if j.get("source_platform") == p]
        stats[p] = {
            "fetched": len([j for j in all_jobs_raw  if j.get("source_platform") == p]),
            "ok":      len([j for j in platform_jobs if j.get("has_jd")]),
            "fail":    len([j for j in platform_jobs if not j.get("has_jd")]),
        }

    row = [
        now.strftime("%Y-%m-%d"),
        now.strftime("%H:%M"),
        category,
        stats["linkedin"]["fetched"],    stats["linkedin"]["ok"],    stats["linkedin"]["fail"],
        stats["hirist"]["fetched"],      stats["hirist"]["ok"],      stats["hirist"]["fail"],
        stats["indeed"]["fetched"],      stats["indeed"]["ok"],      stats["indeed"]["fail"],
        stats["google_jobs"]["fetched"], stats["google_jobs"]["ok"], stats["google_jobs"]["fail"],
        stats["naukri"]["fetched"],      stats["naukri"]["ok"],      stats["naukri"]["fail"],
        len(all_jobs_raw),
        len(all_jobs_filtered),
        len(all_jobs_new),
        staged_count,
        notes,
    ]

    if os.path.exists(RUN_LOG_FILE):
        wb = openpyxl.load_workbook(RUN_LOG_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Run Log"

        # Header row
        for col, (h, w) in enumerate(zip(RUN_LOG_HEADERS, RUN_LOG_WIDTHS), 1):
            c           = ws.cell(row=1, column=col, value=h)
            c.font      = Font(bold=True, color=WHITE, size=9)
            c.fill      = _fill(DARK_BLUE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "A2"

    ws.append(row)
    r = ws.max_row

    # Colour the JD Fail cells red if > 0
    fail_cols = [6, 9, 12, 15, 18]  # LinkedIn fail, Hirist fail, Indeed fail, GJobs fail, Naukri fail
    for col in range(1, len(RUN_LOG_HEADERS) + 1):
        cell = ws.cell(r, col)
        cell.font      = Font(size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if col in fail_cols and cell.value and cell.value > 0:
            cell.fill = PatternFill("solid", fgColor="FFD7D7")  # light red
        elif r % 2 == 0:
            cell.fill = _fill(ALT_ROW)

    wb.save(RUN_LOG_FILE)
    print(f"  Run log updated → {RUN_LOG_FILE}")


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*62}")
    print(f"  Job Scraper v2 — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"{'='*62}")

    # Guard: fail loudly if Excel is open — avoids silent sheet-skip
    lock = os.path.join(PIPELINE_DIR, "~$Daily_Jobs_Log.xlsx")
    if os.path.exists(lock):
        print("\n  ERROR: Daily_Jobs_Log.xlsx is open in Excel.")
        print("  Close it first, then re-run scrape.bat.\n")
        try:
            input("  Press ENTER to exit...")
        except EOFError:
            pass
        return

    # Recovery: flush any staged jobs not yet written to the sheet
    unlogged = get_unlogged_staged()
    if unlogged:
        print(f"\n  Recovering {len(unlogged)} staged job(s) not yet in sheet...")
        update_daily_log([j for _, j in unlogged])
        mark_staged_logged([fp for fp, _ in unlogged])
        print()

    # Determine today's category — a dashboard-relayed override wins outright;
    # otherwise show the menu and let a human confirm or override interactively.
    category = check_dashboard_override()
    if category is None:
        computed_idx, same_day = get_today_category()
        category = confirm_category(computed_idx, same_day)
    print(f"\n  Category : {category['name']}")
    print(f"  Query    : {category['query']}\n")
    query = category["query"]

    # ── Stage 1: Fetch ────────────────────────────────────────────────────
    print("Stage 1: Fetching from 5 platforms (6 credits)...")
    all_jobs = []

    all_jobs += fetch_google_site("LinkedIn",    "linkedin.com/jobs/view",  query, "linkedin", pages=2)
    all_jobs += fetch_google_site("Hirist",      "hirist.tech/j",           query, "hirist")
    all_jobs += fetch_google_site("Indeed",      "in.indeed.com/viewjob",   query, "indeed")
    all_jobs += fetch_google_site("Naukri",      "naukri.com/job-listings", query, "naukri")
    all_jobs += fetch_google_jobs(query)

    all_jobs_raw = list(all_jobs)  # snapshot before any filtering
    print(f"\n  Total raw: {len(all_jobs)}\n")

    # ── Stage 2: Google Jobs via filter ───────────────────────────────────
    print("Stage 2: Google Jobs via filter...")
    all_jobs = filter_google_jobs_via(all_jobs)

    # ── Stage 2b: Naukri freshness filter ─────────────────────────────────
    print("\nStage 2b: Naukri freshness filter...")
    all_jobs = filter_naukri_freshness(all_jobs)

    # ── Stage 3: Title blacklist ───────────────────────────────────────────
    print("\nStage 3: Title blacklist...")
    all_jobs = apply_title_blacklist(all_jobs)

    # ── Stage 4: Deduplication ─────────────────────────────────────────────
    print("\nStage 4: Deduplication...")
    processed_ids = load_processed()
    all_jobs = hard_dedup(all_jobs, processed_ids)
    all_jobs = soft_dedup(all_jobs)

    all_jobs_filtered = list(all_jobs)  # snapshot after filters, before dedup

    if not all_jobs:
        print("\nNo new jobs today. Nothing to stage.")
        write_run_log(category["name"], all_jobs_raw, all_jobs_filtered, [], 0, notes="No new jobs")
        return

    # ── Stage 5: JD fetch ─────────────────────────────────────────────────
    all_jobs = enrich_jd(all_jobs)

    # ── Stage 6: Output ────────────────────────────────────────────────────
    print(f"\nStage 6: Staging {len(all_jobs)} job(s)...")
    staged = stage_jobs(all_jobs)

    print("\nUpdating Daily Jobs Log...")
    update_daily_log(all_jobs)
    mark_staged_logged(staged)

    # Persist all new job IDs
    for j in all_jobs:
        processed_ids.add(j.get("job_id", ""))
    save_processed(processed_ids)

    # LinkedIn fail rate note for run log
    li_jobs   = [j for j in all_jobs if j.get("source_platform") == "linkedin"]
    li_fails  = [j for j in li_jobs  if not j.get("has_jd")]
    li_note   = f"LinkedIn 429: {len(li_fails)}" if li_fails else ""
    write_run_log(category["name"], all_jobs_raw, all_jobs_filtered, all_jobs, len(staged), notes=li_note)

    print(f"\n{'='*62}")
    print(f"  Done. {len(staged)} job(s) staged.")
    print(f"  Staged dir : {STAGED_DIR}")
    print(f"  Daily log  : {DAILY_LOG_FILE}")
    print(f"  Run log    : {RUN_LOG_FILE}")
    print(f"  Next step  : run job_runner.py to tailor resumes")
    print(f"{'='*62}\n")


if __name__ == "__main__":
    main()
