---
name: visual_director
description: "[EXPERIMENTAL] The talk's staging owner - the peer of latex_forge. Once the narrative and each slide's function are locked, propose for each slide what to create: substrate-decides-theme, reuse the manuscript figure, embed-vs-bake, one idea per slide - then scaffold the slides and drive the build."
---
# visual_director

**Status: [EXPERIMENTAL]**

## Purpose
The talks-domain twin of [`latex_forge`](canon://skills/latex_forge.md): the Trunk-stage owner for a talk. It runs **after** [`script_editor`](canon://skills/script_editor.md) has locked the spine - the narrative and each beat's function are settled - and proposes, per slide, what to *create*. It fills the `slide_builder` slot named in the [talks axioms](canon://domains/talks/axioms.md), and it inherits the [Slide Design Guide](canon://domains/talks/projects/talk_from_manuscript.md) as its craft.

The drafting gate, exactly as for the manuscript: build a slide **only from a locked beat** whose function is decided. A slide that asserts beyond its beat, or invents a figure the manuscript does not have, is a defect.

## The craft (from the slide design guide)
- **Substrate decides the theme.** A deck is dark when its assets are dark (manim on black), paper-light when its assets are light (REVTeX figures on white). Do not pick a theme by taste and fight the assets into it - this is the highest-leverage staging choice.
- **Reuse the figure, do not recolour it.** Manuscript figures are the talk's figures; the slide's accent colour matches the figure's own data colour. A talk-only variant is derived from the original, never invented.
- **Render order is reveal order.** Vital information last - the punchline equation is the final reveal, not the first. Reserve space before placing.
- **Embed vs bake, per asset.** Bake a finished non-interactive asset (a manuscript figure, a rendered manim clip); embed live when the value *is* interactivity (an explorable phase portrait). The decision is per asset, never per talk.
- **One idea per slide.** A slide carrying two ideas carries neither. KaTeX traps: inline maths dies in raw HTML, captions carrying maths must be markdown siblings, display maths breaks at natural operators.

## The pass
For each locked beat, surface the decisions (author picks): the slide's **function** (intro / equation build / figure / result / recap), the **substrate**, and the **asset** (reuse-figure / bake / embed / new schematic). Then scaffold `slides/sNN-*.md` with the function, the message, the placed equation/figure, and the timing - and drive the build through the existing pipeline (`build_slide_meta.py`, `iterate.py`, or `slidev build`). Inspect the rendered slide (not just a screenshot - confirm KaTeX hydrated and figures are legible at projection size), then iterate.

## Invariants
- **Narrative first.** Do not propose visuals before `script_editor` has locked the beat and its function.
- **Substrate decides the theme.** Never impose a theme the assets fight.
- **Reuse the figure.** The manuscript's figures are the talk's; derive variants, never invent.
- **One idea per slide.** Reveal order is render order - the payoff lands last.
- **Trace to the manuscript.** Every slide's content traces to a beat, which traces to the paper.

## Backing
Forge-by-doing on the next real talk, lifting the slide-design-guide rules into a backing once they prove on a second deck (the [talks project](canon://domains/talks/projects/talk_from_manuscript.md) holds the live state). The pipeline scripts (`build_slide_meta.py`, `iterate.py`) are reused as-is; the judgement - function, substrate, asset per slide - stays here.
