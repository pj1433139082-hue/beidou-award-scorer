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

### One command (recommended)

```bash
python install.py --from github
```

Clones the repo, installs into **both** `~/.agents/skills/` and `~/.qwenworkcn/skills/` (so agent and Qwen
can both use it), installs the dependencies, and runs a self-test.

Other sources:

```bash
python install.py --from zip --zip ./beidou-award-scorer.zip
python install.py --from dir  --dir  ./beidou-award-scorer
python install.py --from github --target D:/skills/beidou-award-scorer   # one location only
python install.py --from github --no-setup                              # skip dependency setup
```

Or just ask an agent: *"install the skill from https://github.com/pj1433139082-hue/beidou-award-scorer and set up its dependencies"*.

### Manual

```bash
cp -r beidou-award-scorer ~/.agents/skills/
# or link it, so edits stay in one place
# Windows
mklink /J "%USERPROFILE%\.agents\skills\beidou-award-scorer" "%CD%"
# macOS / Linux
ln -s "$PWD" ~/.agents/skills/beidou-award-scorer
```

### Dependencies

Only two: `python-docx` (read .docx) and `PyMuPDF` (read .pdf). No network calls, no API keys, no database.

```bash
python scripts/setup.py            # report what is missing
python scripts/setup.py --install  # install what is missing, then self-test (extract + check)
```

The health check itself uses only the standard library, so it runs even when the dependencies are absent.

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

## Keeping a private copy in sync

A common setup: this public repo holds the generic rubric, while a private copy holds the real
competition criteria. Both should run **identical scripts**.

```bash
# after changing anything in scripts/, references/ or assets/judgments-template.md
python scripts/sync_check.py /path/to/private-copy
```

Exit 0 = the two copies agree. Exit 1 lists exactly which shared files diverged. The rubric file is
*not* compared — that is the one thing the two copies are supposed to differ on (private copy:
`assets/rubric.json`; public repo: `assets/rubric.example.json`). The scripts pick whichever exists,
so the same code runs in both places without a flag.

CI runs this check too, and asserts it **fails** on an intentionally drifted mirror — a drift detector
that never fires is worse than none.

## What the checks can and cannot do

They verify **form, not truth**. They can say a number is present, consistent, and accompanied by evidence — never that it is real.

Fixed mis-detections worth knowing about: merged cells read as empty, sign direction on percentage drops, percentages picked from a neighbouring row, table headers counted as data, over-broad keyword matches, and award-specific table layouts checked against the generic template.

**Confirm every `fail` against the source before quoting it.**

## Documentation

- [使用说明 (HTML)](docs/使用说明.html) · [PDF](docs/使用说明.pdf) — install, use, outputs
- [流程图：安装 → 使用 → 输出](docs/流程图-安装到输出.md) — one-page visual walkthrough

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