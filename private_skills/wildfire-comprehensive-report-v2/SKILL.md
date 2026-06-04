---
name: wildfire-comprehensive-report-v2
description: "Render the v2 comprehensive submission-review PDF for a wildfire risk modeling submission. Combines (a) the Review Agent's Science + Practitioner rubric reviews, (b) human evaluator commentary from the Preliminary-Extract Science and Practitioner evaluator-feedback xlsx files, organized per-question so every evaluator's answer appears beneath the same rubric question, and (c) a model at-a-glance card. Each report is a multi-page PDF with: cover + source materials block, at-a-glance, evaluator commentary (science + practitioner) listed question-by-question, the Review Agent's full Science + Practitioner rubric review, and a coverage notes section (evidence gaps + branching skips). Choose this skill when the user asks for v2 / three-voice / comprehensive submission-review PDFs, or when the request mentions integrating evaluator commentary with rubric reviews."
---

# wildfire-comprehensive-report-v2 — render the per-submission v2 PDF

## Success criterion — read this FIRST

**The deliverable of this skill is one PDF file on disk per workspace processed:**

- `<workspace>/_comprehensive_report_v2_<team_id>_<scenario>.pdf`

If the cell finishes and any expected PDF is not on disk, the skill has failed regardless of how confident the agent's summary sounded.

The only way to produce these PDFs is to call `generate_report_v2(ws_root, ...)` from the `comprehensive_report_v2` module. You MUST:

1. Use `write_file` to save a wrapper `.py` script that calls the module once per workspace.
2. Use `execute` to run the script.
3. Confirm by running `ls <workspace>` and seeing the `_comprehensive_report_v2_*.pdf` file.

You MUST NOT:

- Skip the script-writing step and just narrate findings.
- Generate the PDF yourself with `reportlab` outside this skill — use the helper which is tested and consistent across submissions.
- Claim success without verifying every expected PDF exists on disk.

## Pre-requisites

Every workspace must already contain:

- `_rubric_science.json`   — produced by `wildfire-rubric-review`
- `_rubric_practitioner.json` — produced by `wildfire-rubric-review`

If either JSON is missing the skill raises and the wrapper script exits. Run the rubric-review skill on any missing workspace first, then come back.

Two evaluator xlsx files must be reachable on disk:

- `Preliminary-Extract-Science-Technical-Evaluator-Feedback-*.xlsx`
- `Preliminary-Extract-User-Practitioner-Evaluation-Feedback-*.xlsx`

The skill auto-discovers them by searching the workspace's parent directory and the current working directory. You can also pass explicit paths to `generate_report_v2()` if you keep them elsewhere.

## Sequential execution — DO NOT parallelize

If asked to produce v2 reports for many workspaces, process them sequentially in the main agent thread. Do NOT spawn subagents to parallelize — the rendering itself is fast (seconds), but staying sequential keeps progress predictable and recoverable.

## Mandatory execution flow

For each workspace path in the user's request:

1. **Skip check.** If the workspace already has `_comprehensive_report_v2_<team_id>_<scenario>.pdf` and the prompt does NOT say "refresh"/"regenerate"/"force"/"redo", skip it and print a one-line "already present" note.

2. **Announce.** Print a live progress marker:

   ```python
   print(f"═══ {workspace_name} — Comprehensive v2 PDF ═══")
   ```

3. **Write the wrapper script.** Use `write_file` (mandatory; do not just call `python -c`).

4. **Execute the wrapper script.** Use `execute`. If it raises `FileNotFoundError` because a rubric JSON is missing, READ the error message — it tells the user exactly which rubric needs to be generated first. Pass that information through clearly in your closing summary; do NOT proceed for that workspace.

5. **Verify each PDF exists.** Run `ls <workspace>` and confirm the `_comprehensive_report_v2_*.pdf` file is there.

6. **Print a short Markdown card** under 10 lines per workspace: workspace name, PDF path, file size in MB, headline counts pulled from the rubric JSONs (science answered, practitioner answered, total evidence gaps, number of science + practitioner evaluator responses found).

## Wrapper script — copy this verbatim

```python
import sys, os
# Resolve the skill location regardless of which install path is used.
for candidate in [
    "/home/jovyan/.deepagents/agent/skills/wildfire-comprehensive-report-v2",
    "private_skills/wildfire-comprehensive-report-v2",
]:
    if os.path.isdir(candidate):
        sys.path.insert(0, candidate)
        break

from comprehensive_report_v2 import generate_report_v2

# One call per workspace. Replace the list below with the workspaces the user asked for.
WORKSPACES = [
    "missing_rubrics/Deep_Synthesis__Vanderbilt_WUI_AGNI_NAR_(Forest)",
    # add more here, one per line
]

for ws in WORKSPACES:
    pdf_path = generate_report_v2(ws)
    print(f"Wrote: {pdf_path}")
```

That's it. All the PDF logic, layout, three-voice composition, and team-aliasing live inside `comprehensive_report_v2.py`. The wrapper just passes workspace paths and prints the resulting filenames.

## What the report contains

In order:

1. **Cover** — team name, model name, scenario (Forest/Prairie), generated date, source-materials block (team self-report PDF filename if present in workspace, evaluator response counts).
2. **Model at a glance** — compact 10–13 row table of key attributes pulled from the Review Agent's rubric (Static vs Dynamic, weather variables, ember dispersion, etc.).
3. **Evaluator Commentary — Science Rubric** — per question: the rubric question text, then each science evaluator's answer listed beneath it. Questions no evaluator answered are skipped.
4. **Evaluator Commentary — Practitioner Rubric** — same structure.
5. **Review Agent — Science Rubric Based Evaluation** — every question with the agent's answer + details, or a branching/gap annotation.
6. **Review Agent — Practitioner Rubric Based Evaluation** — same.
7. **Coverage notes from the Review Agent** — evidence_gaps + branching_notes per rubric.

## Output file

Each rendered PDF lands inside the workspace folder as:

```
<workspace>/_comprehensive_report_v2_<team_id>_<scenario>.pdf
```

This keeps the PDF next to its source rubric JSONs and matches the existing `wildfire-comprehensive-report` skill's per-workspace output pattern.
