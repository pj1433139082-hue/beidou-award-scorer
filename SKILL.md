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

## Install

```bash
# one command: clones, installs to both skill roots, prepares dependencies, self-tests
python install.py --from github
# or from a local package / folder
python install.py --from zip --zip ./beidou-award-scorer.zip
python install.py --from dir  --dir  ./beidou-award-scorer
```

`install.py` installs into `~/.agents/skills/` and `~/.qwenworkcn/skills/` (both, so agent and Qwen can use it),
then runs `scripts/setup.py --install` to install the two dependencies and self-test.

Dependencies: `python-docx` (read .docx), `PyMuPDF` (read .pdf). Nothing else, no API keys, no database.

Health check at any time:

```bash
python scripts/setup.py            # report what is missing
python scripts/setup.py --install  # install what is missing, then self-test
```

The health check itself uses only the standard library, so it runs even when the dependencies are absent.

## Before the first run: check dependencies

Run `python scripts/setup.py` once. If it reports a missing dependency, run
`python scripts/setup.py --install` and tell the user what was installed. If `setup.py` cannot run
at all (no Python on PATH), say so plainly and point at the install command — do not proceed to scoring.

## One command, complete output

```bash
python scripts/batch.py "<材料目录>" --out out/ --rubric assets/rubric.example.json --all
```

`--all` = `--with-docx --with-pdf`. One run produces, per submission:

| File | What it is |
|---|---|
| `<name>.doc.json` | extracted text and tables |
| `<name>.checks.json` | the 20 deterministic checks, each with evidence |
| `<name>.score.json` | dimension scores, caps, deductions, disputes |
| `<name>.report.md` | readable report |
| `<name>.docx` | the review note: what is wrong in the original, and how to fix it |
| `<name>.pdf` | same, rendered to PDF (when LibreOffice is available) |

Plus, once per batch:

| File | What it is |
|---|---|
| `leaderboard.md` | ranking |
| `北斗奖评分汇总.xlsx` | **one sheet**: each material is a row (score, bands, dimension cells, projected gain, projected total) with its advice lines and dispute block directly beneath |
| `北斗奖评分意见汇总.docx` | all review notes in one volume (cover + contents) |
| `北斗奖评分意见汇总.pdf` | the same volume as PDF |

PDF needs LibreOffice. Without it the run still succeeds; it prints that PDF was skipped and you
keep the Word files. Point `BD_SOFFICE` at `soffice.exe`, or set `BD_KIT_NODE` and `BD_KIT_CLI`
to use a bundled LibreOffice Kit.

Omit `--all` to get only JSON / Markdown / Excel.

### Verifying the output

Do not trust the log; check the files.

```bash
python scripts/verify_outputs.py out/ --expect-docx --expect-pdf
```

It compares what should exist against what does, and exits 1 if anything is missing. `batch.py`
writes `manifest.json` (every file with its size) and runs this check itself at the end.

**Format is checked too, not just presence:**

```bash
python scripts/verify_format.py out/
```

17 assertions covering: one sheet only, headers present, no blank rows, subtotals and grand total,
no coloured text in Word, four-section structure, three-part advice, PDF count matches docx,
format checks carry no deduction, and the M-numbers in Excel and Word are the same set.
`batch.py` runs it automatically; you do not need to open a single file to know whether the layout
is right.

**And one layer no automated check can cover:**

```bash
python scripts/verify_checks.py out/ --sample 3
```

The three checks above verify code logic, file presence and layout. None of them verifies whether
a machine finding is *correct*. That gap has a real cost: C08 once read `0.50万元` as 50.0 because
its pattern was `\d{2,6}`, and `6.00万元` as 0.0. Nothing failed, nothing warned - the conclusion
was simply wrong, and it stayed wrong until an independent reviewer went back to the source
documents. `verify_checks.py` samples a few submissions and prints every fail/warn next to the
passage it is based on, including a search of the original text, so a human can judge whether the
finding holds. It flags 证据在原文找不到 when a finding's own evidence does not appear in the
document - which is exactly how that C08 bug would have surfaced.

For a pipeline where a missing file must stop the run:

```bash
python scripts/batch.py "<材料目录>" --out out/ --rubric <rubric.json> --all --strict
```

`--strict` exits 2 when any expected artifact is absent, when PDF was requested but no LibreOffice
was found, or when a submission has no `judgments.json`. Without it those conditions are reported
and the run still succeeds.

### What "complete" means per submission

A scored submission must have all of: `doc.json`, `checks.json`, `score.json`, `report.md`,
`docx`. A submission with no `judgments.json` stops after `checks.json` - that is intended, since a
score without human judgement would be invented.


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
| C13 result iteration (award-specific) | warn | treated as "needs confirmation" (mentions iteration but no comparison anchor) — **does not trigger the veto** |
| C13 result iteration (award-specific) | fail | **veto pending review** — **the score is entirely unaffected** (iteration is not a dimension: it carries no points, adds nothing, deducts nothing); only the dispute is pinned to the top and the band is annotated "pending review" |

Fixed deductions apply to the subtotal: D1 duplicate benefit −3, D2 number mismatch −2, D3 missing baseline −2.

**Format problems are flagged, never deducted** — see the section below.

Check IDs and their award scoping live in `scripts/checks.py`; edit that file and the rubric together when changing an award's criteria.

### Format problems are flagged, not deducted

Two checks describe how finished the document is, not how good the work is:

| Check | What it sees | Treatment |
|---|---|---|
| C09 | leftover placeholders (underscores, XXXX, 待填) | dispute only, **no deduction** |
| C17 | no screenshots or images anywhere | dispute only, **no deduction** |

Both still appear in the dispute list, annotated 不扣分, so the submitter knows to fix them.

Why: these are things everyone fixes before submitting — a placeholder in a draft says nothing
about whether the innovation is sound. Deducting points for them does real damage twice over:
it lowers a score for a cosmetic reason, and it pushes a trivial item into the ranked advice list
where it competes with the argument that actually needs strengthening. A submission with twelve
unfilled underscores is not twelve times worse than one with a single placeholder.

The check still runs and still reports, because a reviewer needs to know the document is
unfinished. It just does not move the number.

## How result iteration (C13) is judged

**Every project must supply an iteration statement** — no exemption for first-time entrants.

A `pass` needs two elements within 100 characters of each other:
1. **Comparison anchor** pointing at a past entry: 上一届/上届/往届/历届/过去获奖版本/原版本/原方案/第N届 (N < current). **The current edition does not count.**
2. **Change content**: 迭代/升级/重构/改造/替换/新增/加入/引入/上线, or a technical object followed by 优化/改进/增强/完善/提升 (e.g. "控制算法进行优化").

| Result | Meaning |
|---|---|
| pass | both elements matched |
| warn | change mentioned but no anchor — self-reported, needs human confirmation, **no veto** |
| fail | no change wording at all — iteration cannot be substantiated |

"拓展/扩展" are excluded (they usually mean "promotion scope expanded", not a version change), and
"优化/改进" need a technical object so that "人员工作安排优化" is not mistaken for iteration.

## Why a veto does not zero the score

The rubric marks result-iteration as a veto item, but the same rules also state that the final call
belongs to a review panel (a >10% divergence between human and AI scoring goes to panel review), and that
the AI's job is to *produce a dispute list that drives follow-up questioning* — the source material itself
lists "one-way input/output, no means of re-confirmation" as a known weakness to fix.

So this tool **escalates instead of unilaterally failing a submission**: the dimension is zeroed (that zero *is* the
penalty), the dispute is pinned to the top of the list, and the band is annotated "pending review".
No extra points are invented, because the source material never states a deduction size for a veto item.
Iteration is a gate, not a dimension: whether a submission scores 66 or 48 has nothing to do with it.
Whether the submission is eliminated is decided by the panel after follow-up questioning. To restore a hard veto, change `veto_pending` back to `veto` in `scripts/score.py`.

## Two cases that produce no score

- **Unscored (`score: -1` or omitted)**: the dimension is recorded in `unscored`, **excluded from the denominator**, and the remaining dimensions are rescaled to 100. The report lists "unscored dimensions" separately. Never record "we could not read it" as "the submission is weak" — that is the single most damaging error a scoring tool can make.
- **Invalid rubric bands**: `validate_rubric_bands()` runs at startup and refuses to score if evidence-grade intervals overlap (they are half-open, left-closed). An overlapping rubric is ambiguous and must be fixed, not silently resolved.

## Reading the output

`report.md` has five sections: verdict (score / band / confidence / blended score) → dimension scores (with caps and quote detail) → deductions → **dispute list** → improvement suggestions.

- **Band** (top 85 / second 75 / third 65 / shortlist 55) is a relative scale, not an official cutoff. Without ground-truth ranking data, absolute calibration is impossible.
- **Confidence** drops to medium at ≥6 warn+fail checks or any dimension at grade C/D, and to low at ≥9.
- **Dispute list** is written for the defense panel: it is the "contested points" feedback channel most review processes now ask for.

## Extraction: no docx library, and how exact it is

`extract_report.py` reads `word/document.xml` directly with `zipfile` + `ElementTree`.
It does not import `python-docx`, which spends most of its time building a full DOM tree,
parsing styles.xml / numbering.xml / theme, and wrapping every run in a Python object. All this
tool needs is text, table cells, heading levels and image positions.

Measured on the same twenty submissions: **14.2s with python-docx, 8.1s without.**

The trade-off, stated plainly: the reconstructed `plain` text matches the old extractor
byte-for-byte on 7 of 20 documents. The differences are cover-page tables and nested tables,
where python-docx's cell-merge handling and a direct XML walk disagree by a space or a slash.
**This does not change any conclusion**: the twenty machine checks agree on 19 of 20 documents
(the one exception is C17 counting two images python-docx missed - the new extractor is the more
accurate of the two), and quote verification is identical given the same judgments.

If you ever need the old byte-exact behaviour: `git show 130f4ef:scripts/extract_report.py`.

## Limits of the checks (state these to the user)

- Deterministic checks verify **form, not truth**: they can tell whether a number is present and consistent, never whether it is real.
- Fixed mis-detections worth knowing about: merged cells read as empty, sign direction on percentage drops, percentages picked from a neighbouring row, table headers counted as data, over-broad keyword matches, and award-specific table layouts checked against the generic template. New ones remain possible — **confirm every fail against the source before quoting it.**
- Rubric values are only as good as their source. Keep the citation for each dimension's max score in `references/scoring-method.md` or your own rubric notes.

## How the rubric is resolved

The scripts resolve the default rubric by trying `assets/rubric.json` first, then
`assets/rubric.example.json`. That lets this repo and a private copy share **identical scripts**
while differing only in the rubric file. Pass `--rubric` to be explicit.

Award-specific checks are driven by `profile` in the rubric, not by award names hard-coded in Python.
An award must declare both `profile` and `filename_prefixes`; without them, batch runs cannot detect the
award from a filename and the award-specific checks all fall to `na`.

## References

- `references/scoring-method.md` — evidence grading, cap mapping, quote verification, composition formula, and the reasoning behind them.
- `references/extending.md` — swap the rubric, add an award, add a check, add a file format.
- `assets/rubric.example.json` — a generic rubric demonstrating the schema (dimensions, anchors, gates, deductions, composition rules).
- `assets/judgments-template.md` — the reviewer's input format and scoring discipline.

## Dependencies

Python 3.10+, `python-docx` (docx), `PyMuPDF` (pdf). No network, no external API.