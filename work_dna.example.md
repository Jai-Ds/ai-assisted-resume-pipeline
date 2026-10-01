---
name: Your Work DNA
description: The deep context about your real day-to-day work, wins, and honest weak areas that the tailoring step uses to enrich resumes beyond what's already written in your base resumes.
type: user
---

<!--
This file is read by job_runner.py (path set in config.json → paths.work_dna_file)
and fed into every tailoring prompt. It is NOT a resume — write it informally,
for an AI reading it, not a recruiter. The more honest and specific it is, the
better the tailored output: generic filler here produces generic resumes.

Delete these instructions and the example content below, then write your own.
-->

## Role

Your current title and company, one line.

---

## Day-to-Day Reality

What you actually do, in plain language — the stuff that's true but doesn't
fit neatly into resume-bullet format. Call out anything your resume
undersells: a recent role change, a project that changed what you actually
do day to day, a skill you use constantly but never named explicitly.

---

## Wins

Your real, provable wins — awards, promotions, metrics you can stand behind.
These get anchored on across every tailored resume regardless of JD.

---

## Honest Weak Areas (never overstretch)

List what you genuinely DON'T have, so the tailoring step never invents it.
Rule 1 of every tailoring prompt is "never fabricate" — this section is what
makes that enforceable. Be specific: "no hands-on Kubernetes" is useful,
"not a technical person" is not.

---

## Tools & Tech

Your real day-to-day tools, especially the ones you're deeply hands-on with
vs. ones you've only touched lightly. Note which JD phrases should trigger
surfacing which tool.

---

## Named Skills/Keywords That Must Not Get Lost in Tailoring

List any skill that's genuinely true of your work but tends to live only in
one sentence of your resume's summary or a single bullet, rather than as a
named competency. These get dropped first when a tailoring call rewrites your
summary for a new JD — naming them here protects them.

Also list anything you do NOT want added just because a JD asks for it
(the opposite problem — JDs will ask for things you don't have; name them
here so the model never adds them).

---

## How to Match JD Keywords

A table is the easiest format:

| JD asks for | Draw from |
|-------------|-----------|
| Example: "stakeholder management" | Which real experience backs this? |
| Example: "a skill you're weak on" | How should this be framed honestly, if at all? |

---

## How to Use This in Resume Tailoring

Any standing instructions for the tailoring step — things to always anchor
on, things to never claim, specific phrasing you've deliberately chosen for a
sensitive area of your experience (and why), near-universal keywords for your
field that should never get dropped from Core Competencies just to make room
for a niche JD-specific term.
