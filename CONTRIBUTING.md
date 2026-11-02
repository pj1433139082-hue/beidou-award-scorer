# Contributing

## What belongs here

Fixes and additions to the **scoring engine**: extraction, deterministic checks, evidence grading, quote verification, report rendering, and the rubric schema.

What does **not** belong here:

- Any real competition's confidential scoring criteria. The shipped rubric is a generic example. If you need a private rubric, keep it outside this repo and pass it with `--rubric`.
- Submission documents. Never commit real submissions, or any derived `.doc.json` / `.checks.json` / `.judgments.json` (they are git-ignored for this reason).
- LLM API integration. The design deliberately keeps scoring deterministic and offline; see `references/scoring-method.md` §8.

## Adding a check

1. Write `def cNN_name(doc, award, rub): ... return R(...)` in `scripts/checks.py`.
2. Append it to the `CHECKS` list.
3. If it should only run for some awards, gate it on `profile(award, rub).get("your_flag")` and declare that flag in the rubric's `awards.<name>.profile`.
4. Return `R(...)` with `dims` set to the affected dimension ids and `evidence` pointing back to the source (table N / row M, or a text fragment).
5. If you cannot decide, return `status="na"` — do not force a verdict to make the check "produce output".

## Adding a test fixture

There is no test suite yet. The minimum bar for a PR that changes detection logic:

- Run the changed check against a synthetic document that triggers it.
- Run it against one that must **not** trigger it.
- Paste both outputs in the PR description.

Known false-positive classes are listed in the README; a PR that adds a new one should add it there too.

## Style

- Python: standard library + `python-docx` + `PyMuPDF` only. No new dependencies without discussion.
- No network calls, no telemetry, no timestamps in output — determinism is a tested property.
- Chinese and English are both fine in code comments; keep user-facing strings in the language of the document being scored.

## Commit messages

`<area>: <what changed>` — e.g. `checks: stop counting table headers as data rows`.
