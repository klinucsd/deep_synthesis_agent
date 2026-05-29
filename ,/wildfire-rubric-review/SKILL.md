---
name: wildfire-rubric-review
description: Generate BOTH wildfire rubric reviews (Science Evaluator AND Practitioner Evaluator) for a submission in one cell. DEFAULT choice for any request mentioning "rubric reviews" (plural), "both rubrics", "wildfire review", or unqualified review requests. Reads each workspace file ONCE then builds both rubric JSON files from the shared evidence — faster and cleaner than running the two single-rubric skills back-to-back. Prefer this skill over wildfire-rubric-science or wildfire-rubric-practitioner unless the user explicitly asks for only one rubric or for the filled .xlsx output.
---

# wildfire-rubric-review — both rubrics, one pass

## Success criterion — read this FIRST

**The deliverable of this skill is two files on disk:**

- `<workspace>/_rubric_science.json`
- `<workspace>/_rubric_practitioner.json`

If the cell finishes and these two files do not exist, the skill has
failed regardless of how confident the agent's summary sounded.

This skill **does NOT generate the filled xlsx** — the boss's web forms
have replaced the Excel-upload reviewer workflow. If a reviewer needs
the filled xlsx for some reason, run `wildfire-rubric-science` and
`wildfire-rubric-practitioner` separately (they still produce xlsx).

The only way to produce the two JSON files is to call
`write_rubric_review(..., write_xlsx=False)` from the `wildfire_rubric`
module **twice** (once for each rubric type). You MUST:

1. Use `write_file` to save ONE wrapper `.py` script that does all three
   steps below.
2. Use `execute` to run that script with
   `python /path/to/wrapper_script.py`.
3. Confirm by running `ls <workspace>` and seeing both JSON files appear.

Specifically, you MUST NOT:

- Skip the script-writing step and just narrate findings.
- Read each workspace file twice (once for science, once for practitioner)
  — defeats the purpose of this combined skill.
- Generate the JSON yourself with `json.dumps(...)` — that path bypasses
  the validator.
- Claim success without verifying both JSON files exist on disk.

## What this skill does

Reads a downloaded workspace from the Wildfire Risk Modeling Exercise
project ONCE and produces both rubric reviews:

- **Science Evaluator** rubric — 65 questions: 59 from the
  Excel rubric + 6 from the Science web form (ids `qf1`–`qf6`).
- **Practitioner Evaluator** rubric — 53 questions: 38 from the
  Excel rubric + 15 from the Practitioner web form (ids `qf1`–`qf15`).

Both JSONs include the same four buckets — `answers`, `branching_notes`,
`evidence_gaps`, `team_only`. The `team_only` bucket is auto-populated
for questions marked as subjective ratings (the review agent does not
extract these from materials).

## Pre-requisites

- Workspace must already be downloaded by `ndp-projects` /
  `ndp-workspaces` (a `_manifest.json` file in the workspace folder is
  the marker).
- This skill is **self-contained** — it ships its own copy of the helper
  module and the rubric template:
  `~/.deepagents/agent/skills/wildfire-rubric-review/wildfire_rubric.py`
  and `Exercise-Evaluation-Rubric-20260518.xlsx`. No other skills need to
  be installed. (The single-rubric skills `wildfire-rubric-science` and
  `wildfire-rubric-practitioner` exist alongside but are only needed for
  the single-rubric or .xlsx-output workflows; this skill works on its own.)

## Sequential execution — DO NOT parallelize

Reviews must run sequentially in the main agent thread. Do NOT spawn
subagents to evaluate multiple workspaces in parallel; the LLM rate
limit makes parallel review unworkable at the 30-submission scale.

## Mandatory execution flow — three steps in ONE wrapper script

Write a single `_rubric_review_<workspace>.py` wrapper that performs
all three steps below in order. The agent's job is to (a) write that
script, (b) execute it, (c) verify outputs.

### Step 0: Print both rubric schemas (sanity)

Before answering anything, write and run a short script that imports
both `SCIENCE_RUBRIC` and `PRACTITIONER_RUBRIC` and prints every entry's
`id`, `text`, `type`, `options`, and `branch_when`. Keep that schema
dump in scrollback so each subsequent answer is grounded in the actual
question wording. This is mandatory because past runs have produced
answers keyed to the wrong question.

```python
import sys
sys.path.insert(0, "/home/jovyan/.deepagents/agent/skills/wildfire-rubric-review")
from wildfire_rubric import SCIENCE_RUBRIC, PRACTITIONER_RUBRIC

for label, schema in [("SCIENCE", SCIENCE_RUBRIC), ("PRACTITIONER", PRACTITIONER_RUBRIC)]:
    print(f"\n=== {label} RUBRIC ({len(schema)} questions) ===")
    for q in schema:
        print(f"{q['id']}: type={q['type']}  branch={q.get('branch_when')}")
        print(f"  Q: {q['text']}")
        if q.get("options"):
            print(f"  Options: {q['options']}")
```

This can be a separate small script (not the wrapper) — just so you see
the schemas before writing the wrapper.

### Step 1: Read the workspace ONCE (the wrapper's first block)

The same 3-layer file-selection strategy from
`wildfire-rubric-science/SKILL.md` applies. Read everything ONCE into
kernel variables and reuse them for both rubrics.

**PDF reading — use the `.txt` sidecar pattern, never print content**

For each PDF in the workspace:

A. **Extract to a `.txt` sidecar** with a tiny script (length-only output):

```python
import sys
from pathlib import Path
sys.path.insert(0, "/home/jovyan/.deepagents/agent/skills/wildfire-rubric-review")
from wildfire_rubric import extract_pdf_text

pdf_path = Path("/home/jovyan/.../FinalReport.pdf")
txt_path = pdf_path.with_suffix(".pdf.txt")
if not txt_path.exists():
    txt_path.write_text(extract_pdf_text(str(pdf_path)))
print(f"Extracted {txt_path.stat().st_size:,} bytes → {txt_path.name}")
```

B. **Then use the `read_file` tool on the `.txt` sidecar.** That puts
the PDF text into your reasoning context without rendering it in the
cell (Sage collapses `read_file` results under "details" by default).

**CRITICAL — never print PDF content to stdout.** No
`print(pdf_text)`, no `print(pdf_text[:2000])`, no passing full PDF text
through `_sage_progress`. The cell display becomes unusable when 20-page
PDFs are dumped. Length-only confirmations (`print(f"Extracted N bytes")`)
are fine; substantive content is fine only via `read_file` on the .txt
sidecar (which Sage hides in the cell but exposes to the agent).

The full wrapper script (Step 2/3) reads READMEs, notebooks, configs,
etc. directly. Only PDFs need the txt-sidecar detour.

### Step 2: Build the Science rubric answers and write JSON

Iterate `SCIENCE_RUBRIC` (65 questions) and produce an `answers` dict.
For each question, decide whether the question applies and, if so,
extract the answer from the kernel variables loaded in Step 1.

Same answer shapes as the single-rubric skills:
- `single_select` → `{"value": <one option>, "notes": "..."}`
- `multi_select`  → `{"selected": [<options>], "notes": "..."}`
- `narrative`     → `{"notes": "..."}`
- `text` (q1, q2) → `{"value": "<string>", "notes": ""}`
- Not applicable due to branching → omit and record in `branching_notes`.
- Could have applied but evidence missing → omit and record in `evidence_gaps`.
- `team_only` questions (where `spec.get("team_only") == True`) → set
  `answers[qid] = None` and do nothing else. The `team_only` bucket is
  auto-populated by `write_rubric_review()`.

```python
# Step 2 — Science rubric
science_answers = {}
science_branching = {}
science_gaps = {}

# ... iterate SCIENCE_RUBRIC, fill answers / branching / gaps ...

write_rubric_review(
    ws_root,
    rubric_type="science",
    answers=science_answers,
    branching_notes=science_branching,
    evidence_gaps=science_gaps,
    write_xlsx=False,          # ← combined skill skips xlsx
)
print(f"Wrote _rubric_science.json: "
      f"{sum(1 for v in science_answers.values() if v is not None)} answered, "
      f"{len(science_branching)} branching, {len(science_gaps)} gaps")
```

### Step 3: Build the Practitioner rubric answers and write JSON

Repeat for the Practitioner schema (53 questions). Same answer shapes.

```python
# Step 3 — Practitioner rubric
practitioner_answers = {}
practitioner_branching = {}
practitioner_gaps = {}

# ... iterate PRACTITIONER_RUBRIC, fill answers / branching / gaps ...

write_rubric_review(
    ws_root,
    rubric_type="practitioner",
    answers=practitioner_answers,
    branching_notes=practitioner_branching,
    evidence_gaps=practitioner_gaps,
    write_xlsx=False,
)
print(f"Wrote _rubric_practitioner.json: "
      f"{sum(1 for v in practitioner_answers.values() if v is not None)} answered, "
      f"{len(practitioner_branching)} branching, {len(practitioner_gaps)} gaps")
```

## Verify before reporting success

Inside the wrapper script, add these assertions right before each
`write_rubric_review(...)` call (for both rubrics):

```python
filled = {k for k, v in answers.items() if v is not None}
assert not (filled & branching_notes.keys()), \
    f"BUG: same id in answers and branching_notes: {filled & branching_notes.keys()}"
assert not (filled & evidence_gaps.keys()), \
    f"BUG: same id in answers and evidence_gaps: {filled & evidence_gaps.keys()}"
```

If either assertion fires, fix the conflict and re-run. Do not bypass.

After the wrapper script completes, run `ls <workspace>` and confirm
both `_rubric_science.json` AND `_rubric_practitioner.json` exist. Then
print a short Markdown summary card with:
- workspace name
- science: answered / branching / gaps / team_only counts
- practitioner: answered / branching / gaps / team_only counts

## Cross-rubric tip — Practitioner cross-references Science

Some Practitioner questions cross-reference findings the Science rubric
already established (e.g., Practitioner q8 = "static or dynamic?" maps
to Science q4; Practitioner q3/q4 primary/secondary use map to Science
q20/q20a). Since Step 2 has already produced the science answers dict
IN-MEMORY, the Practitioner step can reference it directly:

```python
# In Step 3 — reuse science answers
static_or_dynamic = science_answers["q4"]["value"]  # already computed
practitioner_answers["q8"] = {"value": static_or_dynamic, "notes": "..."}
```

This avoids re-deriving the same fact from the PDF a second time.

## Epistemic discipline — what you can and can't infer

The same rules from `wildfire-rubric-science/SKILL.md` apply:
- For METHODOLOGY questions, absence of evidence is not evidence of
  absence — use `evidence_gaps` for these.
- For ARTIFACT questions (e.g., "does the team's CSV have column X?"),
  absence IS evidence — answer "No" with a note.
- Each `notes` / `evidence_gaps` / `branching_notes` entry must begin
  with a paraphrase of the actual question wording.

## Wrapper script template — copy and adapt

```python
"""Combined rubric review for one workspace. Produces:
   _rubric_science.json   (65 questions)
   _rubric_practitioner.json (53 questions)
No xlsx — the boss's web forms replace the filled-rubric reviewer flow."""
import json, sys
from pathlib import Path

sys.path.insert(0, "/home/jovyan/.deepagents/agent/skills/wildfire-rubric-review")
from wildfire_rubric import (
    SCIENCE_RUBRIC, PRACTITIONER_RUBRIC,
    extract_pdf_text, write_rubric_review,
    branch_skip, cascade_skip,
)

ws_root = Path("<ABSOLUTE WORKSPACE PATH HERE>")

# ====================================================================
# Step 1: read workspace ONCE
# ====================================================================
manifest = json.loads((ws_root / "_manifest.json").read_text())
pdf_text = extract_pdf_text(
    f"{ws_root}/additional_resources/<...>/<FinalReport>.pdf"
)
# ... add reads of README, notebooks, configs as needed ...

# ====================================================================
# Step 2: Science rubric
# ====================================================================
science_answers = {
    # "q1": {"value": "<model name>", "notes": ""},
    # "q3": {"value": "Yes" / "No", "notes": "..."},
    # ...
    # "qf2": None,   # team_only — Sage skips, auto-populated
    # "qf5": None,   # team_only
}
science_branching = {
    # "q4a": branch_skip("q4", "Static", "If DYNAMIC - ..."),
    # ...
}
science_gaps = {
    # "qX": "Question: '<paraphrase>' — <why no evidence>",
}

# Sanity assertions
_filled = {k for k, v in science_answers.items() if v is not None}
assert not (_filled & science_branching.keys())
assert not (_filled & science_gaps.keys())

write_rubric_review(
    ws_root, rubric_type="science",
    answers=science_answers,
    branching_notes=science_branching,
    evidence_gaps=science_gaps,
    write_xlsx=False,
)
print(f"Wrote _rubric_science.json")

# ====================================================================
# Step 3: Practitioner rubric (reuse science_answers where relevant)
# ====================================================================
practitioner_answers = {
    # "q1": {"value": science_answers["q1"]["value"], "notes": ""},
    # ...
    # "qf9": None,   # team_only
    # "qf12": None,  # team_only
}
practitioner_branching = {}
practitioner_gaps = {}

_filled = {k for k, v in practitioner_answers.items() if v is not None}
assert not (_filled & practitioner_branching.keys())
assert not (_filled & practitioner_gaps.keys())

write_rubric_review(
    ws_root, rubric_type="practitioner",
    answers=practitioner_answers,
    branching_notes=practitioner_branching,
    evidence_gaps=practitioner_gaps,
    write_xlsx=False,
)
print(f"Wrote _rubric_practitioner.json")
```

## What this skill is NOT

- Not the cross-submission report — that's the
  `wildfire-cross-submission-report` skill.
- Not the per-submission comprehensive report — that's the
  `wildfire-comprehensive-report` skill.
- Not a replacement for the two single-rubric skills — if a reviewer
  needs the filled xlsx for some reason, run those directly.
