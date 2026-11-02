# Security

## Scope

This tool processes **local documents**. It makes no network requests, opens no ports, and stores nothing outside the output directory you pass to it.

## What it does with your data

`extract_report.py` reads a `.docx` / `.pdf` and writes a `.doc.json` containing the document's text and tables. That file is plain text and contains the full content of the source document.

- Output files are written only where `--out` points.
- The default `.gitignore` excludes all generated artifacts, so a batch run inside a clone will not accidentally commit them.
- Nothing is uploaded. There is no telemetry.

## Handling sensitive submissions

If submissions are confidential:

1. Keep them and their generated artifacts outside the repo, or in a git-ignored directory.
2. Treat `*.doc.json`, `*.judgments.json`, and `*.report.md` as **as sensitive as the source document** — they contain verbatim excerpts.
3. Do not commit a rubric that contains confidential scoring criteria.

## Reporting a vulnerability

Open a private security advisory on this repository, or email the maintainer listed on the GitHub profile. Please do not open a public issue for anything that could expose submission content.

## Known limitations

- The deterministic checks are heuristic. A crafted document can make a check return `pass` when the underlying claim is unsupported; the quote-verification layer limits, but does not eliminate, this.
- `score.py` trusts `judgments.json`. If you accept judgments from an untrusted party, review the quote list — this is exactly what the "not verified" section of each report is for.
