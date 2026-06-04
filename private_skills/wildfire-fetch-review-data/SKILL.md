---
name: wildfire-fetch-review-data
description: Fetch the wildfire review data bundle from the Fire_Risk_Review Google Drive folder. Use when the user wants to download the submissions that have no NDP workspaces, or fetch the evaluator feedback Excel files for the Fire Risk Modeling Exercise. Downloads (1) a zipped bundle of synthetic workspaces for the submissions without NDP workspaces and unpacks it into the notebook directory as `missing_rubrics/`, and (2) the two Preliminary-Extract Science + Practitioner evaluator-feedback xlsx files. Uses the reviewer's read-only Google Drive token (.gdrive_token.json). Idempotent — files already on disk are skipped.
---

# wildfire-fetch-review-data — download the missing-workspace submissions and evaluator results

## Success criterion — read this FIRST

**The deliverable of this skill is three artifacts on disk in the notebook's working directory:**

- `missing_rubrics/`  — a folder containing 12 subfolders named `Deep_Synthesis__<team>_(Forest|Prairie)/`, each with a team self-report PDF + a `_manifest.json`.
- `Preliminary-Extract-Science-Technical-Evaluator-Feedback-20260603.xlsx`
- `Preliminary-Extract-User-Practitioner-Evaluation-Feedback-20260603.xlsx`

If the cell finishes and any of these three artifacts is not present, the skill has failed regardless of how confident the agent's summary sounded.

The only way to produce these artifacts is to call `fetch_review_data()` from the `fetch_review_data` module. You MUST:

1. Use `write_file` to save a wrapper `.py` script.
2. Use `execute` to run that script with `python /path/to/wrapper_script.py`.
3. Confirm by running `ls -la` in the notebook directory and seeing the three artifacts.

You MUST NOT:

- Skip the script-writing step and just narrate findings.
- Download files yourself with `curl` / `wget` — the Drive IDs and auth handling are in the helper module.
- Claim success without verifying the three artifacts are on disk.

## What this skill downloads

Three files, all sitting in a Drive folder owned by `fire.risk.review@gmail.com`:

| Artifact | Source | Lands at |
|---|---|---|
| `missing_rubrics.zip` (~635 MB) | Google Drive | unzipped to `./missing_rubrics/` |
| Science Evaluator xlsx (~80 KB) | Google Drive | `./Preliminary-Extract-Science-Technical-Evaluator-Feedback-20260603.xlsx` |
| Practitioner Evaluator xlsx (~75 KB) | Google Drive | `./Preliminary-Extract-User-Practitioner-Evaluation-Feedback-20260603.xlsx` |

The zip is large (PinePeak's team report alone is 322 MB and is duplicated for both Forest and Prairie). Expect the initial download to take a few minutes; subsequent runs skip files already on disk.

## Auth prerequisite

The skill reads the OAuth token from `.gdrive_token.json`. It searches these locations in order:

1. `$GDRIVE_TOKEN` (env var, if set)
2. `/home/jovyan/work/_User-Persistent-Storage_CephBlock_/.gdrive_token.json` (pod persistent storage)
3. `~/.gdrive_token.json`
4. `./.gdrive_token.json` (notebook directory)

The token must be a valid OAuth refresh token authenticated as `fire.risk.review@gmail.com` (the shared review account that owns the Fire_Risk_Review folder). Read-only scope is sufficient. If you get an `invalid_grant` or `RefreshError`, the token is expired and needs to be regenerated via `gdrive_token_setup.py` (out of this skill's scope).

## Wrapper script — copy this verbatim

```python
import sys, os
# The skill module sits next to this SKILL.md. Both common install paths:
for candidate in [
    "/home/jovyan/.deepagents/agent/skills/wildfire-fetch-review-data",
    "private_skills/wildfire-fetch-review-data",
]:
    if os.path.isdir(candidate):
        sys.path.insert(0, candidate)
        break

from fetch_review_data import fetch_review_data

result = fetch_review_data(output_dir=".")
for k, v in result.items():
    print(f"{k}: {v}")
```

That's it. The helper handles auth, download, progress reporting, zip extraction, and skip-if-present behavior. After the script finishes:

1. Run `ls -la missing_rubrics 2>&1 | head -20` — should show 12 `Deep_Synthesis__*` subfolders.
2. Run `ls -la Preliminary-Extract*.xlsx 2>&1` — should show both xlsx files.
3. Print a short markdown card summarizing what was downloaded (file sizes, counts).

## Idempotency

The skill skips downloads when the target file is already present and its size matches the Drive size. To force a re-download, pass `force=True`:

```python
result = fetch_review_data(output_dir=".", force=True)
```

The zip extraction is also conditional — if `missing_rubrics/` already exists and contains all 12 expected subfolders, extraction is skipped.

## What the skill does NOT do

- It does not generate rubric reviews on the downloaded submissions. That is the `wildfire-rubric-review` skill's job; call it separately after this one finishes.
- It does not generate classifications. That is the `wildfire-cross-submission-report` skill's job.
- It does not render PDFs. That is the `wildfire-comprehensive-report-v2` skill's job (when it lands).
- It does not upload anything. The token is read-only by design.
