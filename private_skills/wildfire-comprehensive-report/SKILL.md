---
name: wildfire-comprehensive-report
description: Generate a professional PDF comprehensive review for a wildfire risk modeling submission. Combines the Science Evaluator rubric + Practitioner Evaluator rubric + boss's per-submission synthesis questions into one reviewer-grade artifact.
---

# wildfire-comprehensive-report — render the per-workspace PDF report

## Success criterion — read this FIRST

**The deliverable of this skill is one file on disk:**

- `<workspace>/_comprehensive_report_<workspace>.pdf`

If the cell finishes and this file does not exist, the skill has
failed regardless of how confident the agent's summary sounded.

The only way to produce this file is to call
`generate_report(ws_root)` from the `comprehensive_report` module.
You MUST:

1. Use `write_file` to save a wrapper `.py` script to disk.
2. Use `execute` to run that script with
   `python /path/to/wrapper_script.py`.
3. Confirm by running `ls <workspace>` and seeing the PDF appear.
   If not, investigate the stderr output and re-run; do NOT report
   success.

Specifically, you MUST NOT:

- Skip the script-writing step and just narrate findings.
- Generate the PDF directly with `reportlab` outside this skill —
  use the helper which is tested and consistent across workspaces.
- Claim success without verifying the PDF exists on disk.

## What this skill does

Reads two pre-generated JSON files from a workspace and composes them
into a single professional-grade PDF for reviewers:

- `_rubric_science.json` — produced by `wildfire-rubric-science`
- `_rubric_practitioner.json` — produced by `wildfire-rubric-practitioner`

Output PDF sections:

1. **Cover page** — workspace name, model identifier, location, generated date, headline counts
2. **Executive summary + Model at a Glance** — synthesized paragraph + 12-row at-a-glance table
3. **Section 1: Science Evaluator Rubric** — all 59 questions
4. **Section 2: Practitioner Evaluator Rubric** — all 38 questions
5. **Section 3: Synthesis Questions** — 26 per-submission questions from the boss's reviewer worksheet, with intra-PDF hyperlinks to the source rubric questions
6. **Appendix: File Inventory** — what was downloaded, skipped, errored

## Pre-requisites — BOTH rubrics required

The skill **refuses to run** if either `_rubric_science.json` or
`_rubric_practitioner.json` is missing. An incomplete report is
worse than no report — reviewers might mistake it for complete.

If either JSON is missing, the skill prints exactly what to do:

```
Cannot generate comprehensive report. Required inputs missing:
  ✗ /path/.../_rubric_practitioner.json
    -> Run: %%ask Generate a Practitioner Evaluator rubric review for
            the "Deep Synthesis: ELMFIRE (Forest)" workspace.

Run the missing rubric(s) first, then re-run this cell.
```

Run the missing rubrics in earlier cells, then come back to this one.

## Sequential execution — DO NOT parallelize

If asked to produce comprehensive reports for many workspaces, process
them sequentially in the main agent thread. Do NOT spawn subagents to
parallelize — the LLM rate limit makes parallel work unworkable at the
13-submission scale.

## Mandatory execution flow

For each workspace in scope:

1. **Skip check.** If the workspace already has
   `_comprehensive_report_<workspace>.pdf` and the prompt does NOT say
   "refresh"/"regenerate"/"force"/"redo", skip it.

2. **Announce.** Print a live progress marker via `_sage_progress`:

   ```python
   _sage_progress(f"═══ {workspace_name} — Comprehensive PDF ═══")
   ```

3. **Write the wrapper script.** Use `write_file` (mandatory; do not
   just call `python -c`).

4. **Execute the wrapper script.** Use `execute`. If it raises
   `ValueError` because a rubric JSON is missing, READ the error
   message — it tells the user exactly which rubric needs to be
   generated first. Pass that information through clearly in your
   closing summary; do NOT proceed.

5. **Verify the PDF exists.** Run `ls <workspace>` and confirm the
   `_comprehensive_report_*.pdf` file is there.

6. **Print a short Markdown card** under 10 lines: workspace name,
   PDF path, file size in MB, headline counts pulled from the rubric
   JSONs (science answered, practitioner answered, total questions).

## Wrapper script — copy this verbatim

```python
import sys
sys.path.insert(0, "/home/jovyan/.deepagents/agent/skills/wildfire-comprehensive-report")
from comprehensive_report import generate_report

ws_root = "/path/to/Deep_Synthesis__<TeamName>_(Forest)"
pdf_path = generate_report(ws_root)
print(f"Wrote: {pdf_path}")
```

That's it. All the PDF logic, layout, and styling lives in
`comprehensive_report.py`. The agent's job is just to call
`generate_report()` once per workspace.

## What this skill is NOT

- Not a rubric-generator — it consumes existing rubric JSONs only.
- Not a cross-submission synthesis — that's a separate planned skill
  (`wildfire-cross-submission-report`) that will consume all per-workspace
  PDFs and produce a project-level summary.
- Not an editor of the rubric content — if a rubric answer is wrong,
  fix it by re-running the corresponding rubric skill, not by patching
  the PDF.
