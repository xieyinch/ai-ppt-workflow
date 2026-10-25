# Open-source references used by the PPT workflow

Research date: 2026-10-16

## Adopted

### pptx-from-layouts-skill
Repository: `https://github.com/tristan-mcinnis/pptx-from-layouts-skill`
Local checkout: the template-profiling reference skill

Use its central idea for template-driven generation:

1. profile a PPTX template;
2. express page intent as semantic visual types;
3. map visual types to real slide-master layouts;
4. fill placeholders and validate coverage.

A profile was generated for `templates/neon-cyberpunk-ai.pptx` at:
`templates/profiled/neon-cyberpunk-config.json`.
Its bundled example generated a 5-slide deck with validation score 100/100.

Our adapter is `src/plan_layouts.py`; it creates `*-layout-plan.json` before the
existing PowerPoint COM executor runs. The COM executor remains in place because
it already handles our tested template-specific maps, images, charts, notes, and
render verification.

### GX-Alex/html2pptx
Repository: `https://github.com/GX-Alex/html2pptx`
Local checkout: the HTML-to-PPTX fallback reference

Use as an HTML-to-editable-PPTX fallback for fixed 16:9 WebDeck-style HTML.
It extracts browser-rendered DOM/CSS/SVG and converts suitable text and vector
objects to DrawingML. It is not screenshot-perfect and must be visually checked.

## Configured but not adopted as the core engine

### Presenton
A complete self-hosted application/API with editable PPTX and template import.
It is useful as a product reference, but deploying the whole application would
add a second editor, API surface, and model/runtime stack that duplicates our
local workflow.

### PPTAgent / DeepPresenter
Strong research and autonomous-generation references, but its agent runtime and
free-form renderer are broader than the missing template-layout bridge here.

### Frontend Slides / Lewis / image-first decks
Useful for HTML or visual-image routes, but they overlap with Guizang and the
existing image fallback. They are not used as the native PPTX backend.

## Editability contract

- `native`: editable PowerPoint objects;
- `guizang`: editable browser DOM, not native PowerPoint;
- `huashu` / `html2pptx`: editability depends on HTML constraints;
- `image`: slide-sized image, not element-editable.
