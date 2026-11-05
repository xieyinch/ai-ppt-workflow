---
name: ai-ppt-workflow
description: Build editable PowerPoint decks from real templates.
version: 0.1.0
author: xieyinch (xieyinch), Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [PowerPoint, PPTX, templates, presentations, Office]
    related_skills: []
---

# AI PPT Workflow Skill

Use this skill to build editable PowerPoint decks from a real PPTX or POTX template. It preserves the template's master and layout system, fills text by mapping, replaces images/icons and native chart data, then asks PowerPoint to reopen and export every slide for verification.

Three MIT-licensed templates ship in `templates/` (see `templates/README.md`); start there before looking for external templates. Anything outside `templates/` must be licensed by the caller.

## When to Use

- The user asks for a PPT, classroom presentation, thesis deck, project report, or pitch deck.
- The user asks to fill a supplied PowerPoint template.
- The user wants a reusable template-first PPT production workflow.

Don't use it for PDF-only slides, legacy `.ppt` files, or a deck that must be drawn from scratch without a template.

## Prerequisites

- Windows with Microsoft PowerPoint installed and able to open the selected template.
- Python 3.10+.
- Install dependencies with `terminal(command="python -m pip install -r requirements.txt")`.
- Keep the template and all replacement assets inside the project or pass valid project-relative paths.

## How to Run

From the skill/project root:

```text
terminal(command="python scripts/run_workflow.py --config deck.config.json", timeout=600)
terminal(command="python scripts/qa.py", timeout=60)
```

The configuration must define `template`, `content`, `output`, `render_dir`, `notes`, `sources`, and `quality`.

## Quick Reference

- `src/run_workflow.py` — stable one-command entry for planning, build, QA, rendering, and contact sheet.
- `src/agent_workflow.py` — complete native workflow orchestration.
- `src/make_ppt.py` — multi-backend route dispatcher with explicit artifact contracts.
- `src/plan_layouts.py` — semantic page-job and visual-layout planning.
- `src/build.py` — PowerPoint COM implementation.
- `src/qa.py` — package, content, notes, source, render, and media checks.
- `src/make_contact_sheet.py` — visual review sheet generator.
- `src/export_html.py` — optional browser preview exporter.
- `src/update_charts.py` — native chart-data updater.
- `src/list_shapes.py` — inspect template text slots.
- `scripts/new_deck.py` — scaffold a deck from a bundled template (writes `content.json` + `deck.config.json`).
- `scripts/render_previews.ps1` — export every slide to PNG for review.
- `templates/manifest.json` — bundled template catalog.
- `examples/deck.config.example.json` — minimal configuration.
- `examples/content.example.json` — minimal content map.

## Procedure

1. **Choose the scene.** Define audience, purpose, duration, slide count, and one job per slide. Completion criterion: a slide blueprint exists.
2. **Choose a licensed template.** Preferred order: the user's own template → a bundled template in `templates/` (MIT, run `python scripts/new_deck.py --list`) → an external template whose license permits the intended use. Record the license. Completion criterion: the `.pptx`/`.potx` opens in PowerPoint and its license is recorded.
3. **Inspect the template.** Run `terminal(command="python scripts/list_shapes.py <content-json>")` or inspect it with PowerPoint. Completion criterion: every mapped placeholder is copied exactly from the template.
4. **Research and draft.** Keep verified facts separate from interpretation, and write a speaker script plus sources. Completion criterion: notes and source files exist and contain useful content.
5. **Create the mapping.** Use `map` for text. When the same placeholder occurs in sibling cards, use an array; values are consumed in template order. Use `images` for pictures/icons and `charts` for native chart data. Completion criterion: all intended replacements have a target.
6. **Build.** Run `terminal(command="python scripts/run_workflow.py --config deck.config.json", timeout=600)`. Completion criterion: the result reports `status: PASS`, the expected slide count, zero leftover placeholders, and one PNG per output slide.
7. **Run QA.** Run `terminal(command="python scripts/qa.py", timeout=60)`. Completion criterion: QA reports `status: PASS` for the same output named by `deck.config.json`.
8. **Review visuals.** Review every rendered PNG, not only the cover. Check overlap, clipping, font substitution, unrelated stock graphics, image crop, and readable information density. Completion criterion: every slide has been reviewed and any defect is fixed and rebuilt.
9. **Deliver.** Provide the editable `.pptx`, speaker notes, and source list. Do not include private templates, credentials, browser profiles, generated caches, or personal records in a public repository.

## Bundled Templates

| slug | style | use for |
|---|---|---|
| `inner-chapter` | black & white editorial, blue/orange accents, 59 layouts | consulting, reviews, research readouts |
| `inner-chapter-navy` | white ground, deep navy, 微软雅黑 | 技术方案, 课程汇报, corporate |
| `inner-chapter-ink` | warm paper ground, ink + gold, 思源黑体 CN | 材料/建筑/人文类学术汇报 |

Each folder holds `template.pptx` (empty 59-layout master), `starter.pptx` (9-page deck with
realistic content — the fast path), `previews/` and `LICENSE`. Attribution is in
`templates/ATTRIBUTION.md`; keep it when redistributing.

```bash
python scripts/new_deck.py <name> --template inner-chapter-navy   # starter-based
python scripts/new_deck.py <name> --template inner-chapter --blank  # author from layouts
```

## Content Mapping

Text replacement:

```json
{"src": 1, "map": {"Template title": "New title"}}
```

Repeated placeholders:

```json
{"src": 5, "map": {"Placeholder": ["Card one", "Card two"]}}
```

Image/icon replacement preserves the target bounds:

```json
{"src": 1, "images": [{"kind": "picture", "occurrence": 1, "file": "assets/photo.png"}]}
```

Use `kind: "group"` or `kind: "icon"` for a top-level grouped icon, or use a PowerPoint shape `name` when it is stable.

Native chart replacement:

```json
{"src": 24, "charts": [{"occurrence": 1, "categories": ["A", "B"], "series": {"Series 1": [10, 20]}}]}
```

## Pitfalls

- A text-only map cannot distinguish repeated sibling placeholders; use arrays.
- A template's stock photos or icons remain unless explicitly replaced.
- Image replacement preserves position and size but does not infer a template's original artistic crop or mask.
- PowerPoint COM is required for the build and rendering path; a structural ZIP check alone is not visual verification.
- A successful file write is not enough: reopen the output, run QA, and inspect every rendered page.
- Never publish a user's private template, personal identifiers, local absolute paths, browser data, tokens, or generated deliverables without explicit permission.
- Theme-recolored variants only change the theme color/font scheme; decorative colors hard-coded in the master stay put. Recolor the master if the variant must be complete.
- Any image placed to fill a layout must serve that page's content. Decorative stock images that illustrate nothing are a defect, not a design choice.

## Verification

A verified build has all of the following:

- PowerPoint reopened the output successfully.
- The output slide count equals the content entries.
- No known placeholder remains.
- Every output slide has a PNG render.
- `scripts/qa.py` passes against the same `deck.config.json`.
- Notes and sources are present.
- Every rendered slide was visually reviewed.
