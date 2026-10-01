# AI-Assisted Resume Pipeline

An end-to-end job-search pipeline: **scrape** fresh listings → **tailor** a 1-page resume per job with AI → **analyze** ATS fit before you apply. Built for a Product Manager job search originally, but every role-specific piece — search categories, title filters, job titles/dates, resume content — lives in one config file. Point it at your own base resumes and work history, and it works for any role: sales, hardware, engineering, whatever you're searching for.

## What it does

```
job_scraper.py   →  Daily_Jobs_Log.xlsx   →  job_runner.py        →  job_analysis.py
(5-platform        (review & mark          (tailors a resume per    (ATS keyword match +
 search, dedup,      "Shortlisted")          shortlisted job,         fit score + gap
 freshness filter)                           auto-scaled to 1 page)   analysis)
```

1. **`job_scraper.py`** — searches LinkedIn, Hirist, Indeed, Naukri, and Google Jobs for your configured categories, dedupes across platforms, filters out stale/expired listings, fetches the full job description where the platform allows it, and logs everything to an Excel tracker.
2. You open the Excel file, mark the jobs you want to apply to as `Shortlisted`.
3. **`job_runner.py`** — for each shortlisted job, blends your base resumes + your own "Work DNA" context through an AI tailoring call, then builds a 1-page DOCX/PDF — auto-scaling the font and, if it still doesn't fit, automatically trimming the lowest-priority content until it does. Never ships a 2-page resume silently.
4. **`job_analysis.py`** — scores the tailored resume against the job description (keyword match, fit %, named gaps, improvement ideas) and keeps a master tracker of every application.

## Quickstart

```bash
git clone <this-repo>
cd ai-assisted-resume-pipeline
pip install -r requirements.txt
playwright install chromium   # needed for JD fetching from JS-rendered sites
```

1. **Get a [SerpApi](https://serpapi.com/) key** (free tier works for light use) and create a `.env` file:
   ```
   SERPAPI_KEY=your_key_here
   ```
2. **Install the [Claude Code CLI](https://docs.claude.com/claude-code)** and make sure `claude` is on your PATH and logged in — the tailoring and analysis steps shell out to `claude -p`.
3. **Copy `config.example.json` to `config.json`** and fill in your own:
   - `profile` — your name, contact details, education, and your **fixed roles** (the ones that never change no matter which job you're applying to — title, dates, how many bullets each can have, and whether it's optional to include).
   - `paths` — where your base resumes live, where tailored resumes get saved, and where your Work DNA file is.
   - `base_resume_tags` — a map of short tag → base resume filename (no extension). You can have 1 or 10; the tailoring call picks whichever fits each JD best.
   - `categories` — your own search categories and queries (see [Making it work for any role](#making-it-work-for-any-role) below).
   - `title_blacklist` — words that disqualify a listing from being surfaced at all (seniority levels you don't want, noise terms, etc).
4. **Write your own `work_dna.md`** (copy `work_dna.example.md` to get started, point `paths.work_dna_file` at it). This is the single highest-leverage file in the whole pipeline — it's informal context fed straight into every tailoring call: your real day-to-day work, your honest weak areas (so the AI never fabricates them), and any skill that's true but only ever shows up as one line of prose on your resume (those get dropped first when a resume is rewritten for a new JD — naming them here protects them).
5. **Provide your own base resumes** as `.docx` files, one per tag in `base_resume_tags`, at the path you set in `paths.base_versions_dir`.
6. Run it:
   ```bash
   python job_scraper.py     # fetch + log new jobs
   # ... open Daily_Jobs_Log.xlsx, mark jobs "Shortlisted" ...
   python job_runner.py      # tailor + build resumes for shortlisted jobs
   python job_analysis.py    # ATS score + gap analysis
   ```

## Making it work for any role

Nothing in the code names a specific job title, industry, or person — that all lives in `config.json`. A sales rep repoints `categories` to sales-relevant search queries, lists their own fixed roles and base resumes, and writes a `work_dna.md` about their own quota history and deal types. A hardware engineer does the same with their own categories and technical base resumes. The AI tailoring prompt itself is built dynamically from whatever `config.json` and `work_dna.md` contain — it has no PM-specific language baked in.

The one thing worth tuning per field: `title_blacklist` and your `categories` queries are the levers that decide what counts as noise vs. signal for your specific search — spend a little time getting those right and the rest of the pipeline does the work.

## Notable design choices

- **Mandatory 1-page enforcement**: resumes auto-scale font size via binary search, and if that's still not enough, automatically trim the lowest-priority content (oldest role's weakest bullet, least-relevant project, etc.) until it fits — rather than silently shipping an overflowing resume.
- **Freshness filtering that goes beyond the platform's own "posted recently" filter**: Google's own freshness signals are a crawl/index-date proxy, not a true posting date — this pipeline separately verifies freshness per-platform (e.g. decoding the date encoded in a job board's own URL structure) rather than trusting the search engine's filter alone.
- **Protected competency terms**: near-universal keywords for your field, once named in `work_dna.md`, won't get silently dropped from a tailored resume just to make room for a JD-specific term that only one posting cares about.

## Requirements

- Python 3.9+
- A [SerpApi](https://serpapi.com/) account (job search credits)
- [Claude Code CLI](https://docs.claude.com/claude-code), authenticated (`claude -p` is used for tailoring + ATS analysis)
- `playwright install chromium` for JD fetching on JS-rendered job sites
