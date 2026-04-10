---
name: mmis-dev-knowledge
description: Reusable knowledge base for MMIS automation engineering. Use before developing or refactoring any MMIS-related skill, script, login flow, query flow, download flow, or Excel post-processing flow to read prior experience, stable patterns, failure modes, and performance guidance. Use again after finishing MMIS development or fixing MMIS bugs to write back refined knowledge, merge duplicates, and update best practices.
---

# MMIS Dev Knowledge

This skill is a knowledge base, not a task executor.

Use it in two moments:

1. Before MMIS automation development
   - Read the current knowledge base first.
   - Reuse stable patterns instead of rediscovering them.
2. After MMIS automation development or bug fixing
   - Write back reusable lessons.
   - Refine or replace outdated guidance.

## Read Workflow

Start with:

- [knowledge-base.json](references/knowledge-base.json)
- [mmis-web-data-location-guide.md](references/mmis-web-data-location-guide.md) when the task involves Playwright, selector design, DOM inspection, Maximo table scraping, or browser-side MMIS debugging.
- [mmis-query-workflow-and-debug-guide.md](references/mmis-query-workflow-and-debug-guide.md) when the task involves MMIS query execution rules, filter-row input behavior, Enter-triggered filtering, date formatting, batch query stabilization, or debugging why a query returns the wrong result set.

Focus on these top-level sections:

- `loginStrategies`
- `apiPatterns`
- `commonErrors`
- `performanceTips`
- `stableWorkflows`
- `toolingNotes`

Use the guide file for:

- Chrome DevTools inspection workflow
- MMIS / Maximo DOM structure analysis
- selector stability rules
- table row / column locating patterns
- distinguishing empty values from missing data
- Playwright locator and wait strategy

Use the query workflow guide for:

- query trigger rules such as focus-before-Enter
- switching to `所有記錄` before filtering
- field mapping and value-format rules
- common MMIS query failures and their fixes
- batch query debug strategy with logs and screenshots

When planning a new MMIS skill:

1. Read the relevant sections only.
2. Summarize the most relevant guidance for the current task.
3. Prefer the highest-confidence patterns first.
4. Treat `status=confirmed` as the default baseline.
5. Treat `status=experimental` as candidate guidance that still needs validation.

## Write Workflow

When a new MMIS development task finishes:

1. Identify reusable knowledge, not one-off steps.
2. Update the knowledge base with:
   - problem
   - solution
   - status
   - reuse potential
   - cautions
3. Merge with existing entries when the topic already exists.
4. Replace outdated or disproven guidance instead of leaving contradictions.

Use the update script:

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-dev-knowledge\scripts\update_mmis_dev_knowledge.py --input <json-file>
```

The input JSON should contain:

```json
{
  "section": "commonErrors",
  "id": "maximo-shared-session-event-rejection",
  "entry": {
    "title": "Maximo event requests can trigger shared-session rejection",
    "description": "Direct requests to maximo.jsp may be rejected even after successful HTTP login.",
    "applicable_when": [
      "Using requests.Session to replay Maximo UI events"
    ],
    "advantages": [
      "Documents a recurring failure mode early"
    ],
    "disadvantages": [
      "Requires fallback or deeper protocol analysis"
    ],
    "recommended_when": [
      "Designing request-driven MMIS automations"
    ],
    "status": "confirmed",
    "reusable": true,
    "notes": [
      "Login can still be request-driven even if event replay is not fully solved."
    ],
    "replaces": []
  }
}
```

## Update Rules

- Keep entries concise and reusable.
- Prefer editing an existing entry over adding near-duplicates.
- Mark guidance with one of:
  - `confirmed`
  - `experimental`
  - `deprecated`
- If an older entry is wrong, mark it `deprecated` or replace it through `replaces`.
- Keep best practices stronger than anecdotes.

## Expected Output When Reading

When another MMIS skill uses this knowledge base, return:

1. The most relevant guidance for the task.
2. Known risks and likely failure points.
3. The preferred implementation path.
4. Whether browser fallback should be avoided, preferred, or kept as contingency.

## Expected Output When Writing

After updating the knowledge base, report:

1. Which section was updated.
2. Whether the entry was added, merged, or replaced.
3. Any deprecated guidance that was superseded.
4. The path to the updated knowledge file.
