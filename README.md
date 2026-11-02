# beidou-award-scorer

An agent skill that turns an award submission document into a **re-checkable score**.

Deterministic checks compute what can be computed. A reviewer supplies the rest — and every dimension score must carry a **verbatim quote** from the source text. Quotes are matched against the extracted document; an unmatched quote is rejected and the dimension's evidence grade is capped. The output is a report you can defend line by line, plus a dispute list for the Q&A panel.

```
document.docx / .pdf
  → doc.json        ordered text + tables (with section paths)
  → checks.json     20 deterministic checks, each with an evidence locator
  → score.json      weighted dimensions, caps, deductions, disputes
  → report.md       verdict → dimensions → deductions → disputes → fixes
```

## Why it exists

Most "AI grading" tools produce a number nobody can audit. This one is built the other way round:

1. **A technology type is not a bonus.** Whether the work uses AI, digitisation, or conventional engineering is classification only.
2. **No quote, no score.** Every dimension judgement cites verbatim source text (≥6 chars). `score.py` verifies each one.
3. **Implemented ≠ planned.** Quantified benefits must be split into actual and target; mixed or unlabelled benefits count as planned.
4. **"We could not read it" is never "the submission is weak."** Unscorable dimensions are excluded from the denominator, never counted as zero.

## Install

The skill is a plain directory of scripts plus a SKILL.md. Copy or link it into your agent's skill folder:

```bash
# agent-global
cp -r beidou-award-scorer ~/.agents/skills/

# or link it, so edits stay in one place
# Windows
mklink /J "%USERPROFILE%\.agents\skills\beidou-award-scorer" "%CD%"
# macOS / Linux
ln -s "$PWD" ~/.agents/skills/beidou-award-scorer
```

Dependencies: Python 3.10+, `python-docx`, `PyMuPDF`. No network, no API keys, no database.

```bash
pip install python-docx pymupdf
```

## Quick start

```bash
# 1. batch run (extract + check everything in a folder)
python scripts/batch.py "<documents-dir>" --out out/

# 2. single document
python scripts/extract_report.py report.docx --out out/
python scripts/checks.py out/report.doc.json --award 示例奖A --out out/report.checks.json

# 3. read checks.json, then write out/report.judgments.json
#    per assets/judgments-template.md (score / grade / rationale / quotes)

# 4. score
python scripts/score.py out/report.doc.json \
    --checks out/report.checks.json \
    --judgments out/report.judgments.json \
    --award 示例奖A --out out/

# optional: blend with an on-site score (75/25) and flag >10% divergence
python scripts/score.py ... --onsite 78
```

## Bring your own rubric

`assets/rubric.example.json` is a **generic** rubric that demonstrates the schema. It does not correspond to any real competition. Replace it with your own:

```bash
python scripts/batch.py "<docs>" --out out/ --rubric path/to/my-rubric.json
```

A rubric defines, for each award:

| field | meaning |
|---|---|
| `dimensions[]` | `id`, `name`, `max`, official wording, scoring anchors, required evidence |
| `profile` | which award-specific checks to enable (`promotion`, `model_lab`, `product`, `strategy`, `template_sections`, `roi`, `generic_tables`) |
| `filename_prefixes` | how to detect the award from a filename during batch runs |
| `min_chars` | reference length floor |

Global rules live at the top level: `evidence_grades` (A/B/C/D with non-overlapping half-open bands and caps), `gates`, `deductions`, and `scoring_rules` (weighting, aggregation, tie-break, `unscored_sentinel`).

Sum of `max` across dimensions should equal the award's total. `validate_rubric_bands()` refuses to score if evidence-grade bands overlap.

## What the checks can and cannot do

They verify **form, not truth**. They can say a number is present, consistent, and accompanied by evidence — never that it is real.

Fixed mis-detections worth knowing about: merged cells read as empty, sign direction on percentage drops, percentages picked from a neighbouring row, table headers counted as data, over-broad keyword matches, and award-specific table layouts checked against the generic template.

**Confirm every `fail` against the source before quoting it.**

## Repository layout

```
SKILL.md                      what the agent loads
scripts/extract_report.py     docx/pdf → doc.json
scripts/checks.py             20 deterministic checks → checks.json
scripts/score.py              judgments + checks → score.json / report.md
scripts/batch.py              folder → leaderboard.md
assets/rubric.example.json    generic rubric (replace it)
assets/judgments-template.md  reviewer input format + scoring discipline
references/scoring-method.md  evidence grading, caps, quote verification
references/extending.md       swap rubric / add award / add check / add format
```

## License

MIT — see [LICENSE](LICENSE).
