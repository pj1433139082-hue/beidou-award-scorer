---
name: beidou-award-scorer
description: Score innovation-award submission documents (docx/pdf) against a weighted, evidence-graded rubric. Extracts text and tables, runs 20 deterministic checks, requires a verbatim quote for every dimension judgement, caps scores when evidence is missing, and emits a traceable report with a dispute list for defense Q&A. Use when scoring award applications, self-auditing a submission before the deadline, ranking a batch of proposals, or preparing reviewer questions. 触发：评分、申报书打分、评审报告、评分卡、争议点清单.
---

# Beidou Award Scorer

Turn a submission document into a **re-checkable score**: deterministic checks compute what can be computed, a reviewer supplies the rest, and every point must carry a verbatim quote from the source text.

The rubric shipped in `assets/rubric.example.json` demonstrates the structure with a generic five-dimension scheme. Point `--rubric` at your own rubric to score against a specific award's official criteria.

## Three rules that must not break

1. **A technology type is not a bonus.** Whether the work uses AI, digitisation, or conventional engineering is classification only. Adding points because "it uses AI" is the exact bias that most award rubrics try to correct.
2. **No quote, no score.** Every dimension judgement carries `quotes` — verbatim fragments from the source, at least 6 characters. `score.py` matches each one against the extracted text; an unmatched quote is rejected and the dimension's evidence grade is capped.
3. **Implemented ≠ planned.** Quantified benefits must be separated into "actual value (implemented)" and "target (planned)". A benefit that does not distinguish them counts as planned.

## Workflow

```
document.docx/.pdf
   │  scripts/extract_report.py
   ▼
doc.json        ordered body text + tables (with section paths)
   │  scripts/checks.py
   ▼
checks.json     20 checks: pass / warn / fail / na, each with an evidence locator
   │  + judgments.json  (written by the reviewer, with verbatim quotes)
   ▼
scripts/score.py
   ▼
score.json + report.md
```

Batch run with a leaderboard:

```bash
python scripts/batch.py "<documents-dir>" --out out/
python scripts/batch.py "<documents-dir>" --out out/ --award 金刚钻
```

Single document:

```bash
python scripts/extract_report.py report.docx --out out/
python scripts/checks.py out/report.doc.json --award 金刚钻 --out out/report.checks.json
# read checks.json, then write out/report.judgments.json per assets/judgments-template.md
python scripts/score.py out/report.doc.json --checks out/report.checks.json \
    --judgments out/report.judgments.json --award 金刚钻 --out out/
# add --onsite 78 for the 75/25 blended score and the >10% conflict flag
```

Award names are whatever keys your rubric defines; `--award` must match one of them.

## Step 3 — the reviewer's job

1. Read `checks.json` first: the warn/fail lines already say where the document is empty or inconsistent.
2. Read `doc.json`: body and tables are both there; use the `section` field to locate a chapter.
3. Fill `score`, `grade`, `rationale`, `quotes` per dimension.
   - `score` may not exceed **max × that grade's cap** (A 1.00 / B 0.75 / C 0.50 / D 0.20).
   - `quotes` are copied verbatim — no rewriting, no joining, no translation.
4. Anything uncertain goes into `disputes`. **Never adjust a score to hide a disagreement.**

## How scores get capped

`apply_machine()` in `scripts/score.py` maps check outcomes onto dimension ceilings; the strictest applicable ceiling wins.

| Check | Outcome | Effect |
|---|---|---|
| C02 implemented/planned separation | fail / warn | value capped at 0.40 / 0.75 |
| C04 benefit evidence coverage | fail / warn | value capped at 0.70 / 0.85 |
| C11 lab & trial-plant evidence | fail / warn | intelligence capped at 0.40 / 0.70 |
| C12 product readiness | warn / fail | readiness capped at 0.65 / 0.45 |
| C14 novelty contrast | fail / warn | novelty capped at 0.60 / 0.80 |
| C13 result iteration (award-specific) | fail | **veto** — total score forced to 0 |

Fixed deductions (D1 duplicate benefit −3, D2 number mismatch −2, D3 missing baseline −2, D4 missing screenshots −2, leftover placeholders −2) apply to the subtotal.

Check IDs and their award scoping live in `scripts/checks.py`; edit that file and the rubric together when changing an award's criteria.

## Two cases that produce no score

- **Unscored (`score: -1` or omitted)**: the dimension is recorded in `unscored`, **excluded from the denominator**, and the remaining dimensions are rescaled to 100. The report lists "unscored dimensions" separately. Never record "we could not read it" as "the submission is weak" — that is the single most damaging error a scoring tool can make.
- **Invalid rubric bands**: `validate_rubric_bands()` runs at startup and refuses to score if evidence-grade intervals overlap (they are half-open, left-closed). An overlapping rubric is ambiguous and must be fixed, not silently resolved.

## Reading the output

`report.md` has five sections: verdict (score / band / confidence / blended score) → dimension scores (with caps and quote detail) → deductions → **dispute list** → improvement suggestions.

- **Band** (top 85 / second 75 / third 65 / shortlist 55) is a relative scale, not an official cutoff. Without ground-truth ranking data, absolute calibration is impossible.
- **Confidence** drops to medium at ≥6 warn+fail checks or any dimension at grade C/D, and to low at ≥9.
- **Dispute list** is written for the defense panel: it is the "contested points" feedback channel most review processes now ask for.

## Limits of the checks (state these to the user)

- Deterministic checks verify **form, not truth**: they can tell whether a number is present and consistent, never whether it is real.
- Fixed mis-detections worth knowing about: merged cells read as empty, sign direction on percentage drops, percentages picked from a neighbouring row, table headers counted as data, over-broad keyword matches, and award-specific table layouts checked against the generic template. New ones remain possible — **confirm every fail against the source before quoting it.**
- Rubric values are only as good as their source. Keep the citation for each dimension's max score in `references/scoring-method.md` or your own rubric notes.

## References

- `references/scoring-method.md` — evidence grading, cap mapping, quote verification, composition formula, and the reasoning behind them.
- `references/extending.md` — swap the rubric, add an award, add a check, add a file format.
- `assets/rubric.example.json` — a generic rubric demonstrating the schema (dimensions, anchors, gates, deductions, composition rules).
- `assets/judgments-template.md` — the reviewer's input format and scoring discipline.

## Dependencies

Python 3.10+, `python-docx` (docx), `PyMuPDF` (pdf). No network, no external API.
