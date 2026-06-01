---
name: manuscript_spine
description: "[EXPERIMENTAL] Discover the manuscript by interrogating the author - a long, ordered sequence of decisions surfaced from macro to atom (identity, structure, claims, equation/figure placement, literature) that assembles a live, navigable blueprint (the NMD) one node per session, so the final prose pass is mechanical fill-in."
---
# manuscript_spine

**Status: [EXPERIMENTAL]**

## Purpose
The [core-outward trunk](canon://workflow/core_outward_trunk.md) has three stages - Foundry, **Spine**, Trunk. The Foundry rails have owners ([`result_foundry`](canon://skills/result_foundry.md), [`literature_scout`](canon://skills/literature_scout.md), [`numerical_evidence`](canon://skills/numerical_evidence.md)). The Trunk has [`latex_forge`](canon://skills/latex_forge.md). This skill owns the **Spine** - and it owns it by *asking*, not by drafting. The manuscript is in the author's head. The job is to pull it out, one decision at a time, and let the answers assemble the blueprint (the NMD, see [Evidence-Schema Flow](canon://workflow/evidence_schema_flow.md)).

This is the manuscript peer of the talks domain's `script_editor`. It does not write the paper. It *discovers* it - the single most important step, because the blueprint it produces is the drafting instruction set. Get the blueprint right and the prose pass is mechanical.

**The author decides everything. You surface the decisions.** Never invent a thesis, choose a structure, or place a claim by fiat. Each is a question with concrete options the author picks. The narrative and identity docs are a *byproduct* of the answers, not a thing you author up front.

## Self-similar to the macro process
The elicitation is the core-outward trunk in miniature - **orient -> propose one decision -> author decides -> lock -> reconcile -> next**, run at conversation scale. High level first, then propagate down: identity (macro) constrains structure (meso) constrains claims and equation placement (atom). Each level's answers are the substrate the next level is elicited against. You never hold the whole paper in one breath - you grow it node by node, and the framing rings (intro, conclusion, abstract) are elicited **last**.

## The blueprint is a composition manifest, not prose
The blueprint lives at `workflow/state/blueprint.md` in the project root. It holds **commitments and intent, never the sentences**. Its format is line-parseable so [`project_blueprint.py`](canon://skills/manuscript_spine.md) can project it into the live graph after every answer:

```
# Blueprint: <title>
- thesis: <one sentence>
- frame: <the entry frame>
- title: <candidate>
- register: <...>

## <sec-id> | <Section Title> | ring:<core|inner|framing> | status:<stub|blueprinted|drafted>
intent: <what this section must land, and the transition to its neighbour>
- claim: <statement> | result:<Rk> | status:<established|partial|open-gap> [| headline]
- equation: <label> | <latex> | serves:<claim index or text>
- figure: <label> | <caption>
- cite: <citekey> | relation:<supports|contests|context>
```

A claim with no `result:` is a stub, not a locked claim. Duplication is forbidden - a claim stated here may not also be drafted inline (see [Evidence-Schema Flow](canon://workflow/evidence_schema_flow.md)).

## The question ladder (macro -> atom)
Surface one decision at a time. On Claude, use the structured question tool (`AskUserQuestion`); on Codex / Gemini, present a numbered options list and await the choice - client-agnostic, the substrate is the same. After each answer: write it to `blueprint.md`, run the backing to refresh the graph, and let the author see it land in the viewer before the next question.

1. **Identity / frame (macro).** What is the one-sentence thesis? What is the entry frame (the hook the introduction opens on)? What does the paper deliberately *not* cover (scope)? Title candidates? Register? - These answers write the spine block and seed the identity ledger.
2. **Structure (meso).** What are the sections, in reading order? Tag each by ring (core / inner / framing). For each, what must it land, and how does it hand to the next? - writes the section headers and `intent`.
3. **Claims (atom, core-outward).** Take the core sections first. For each section, what claims does it make, in order? Trace each to a result ledger `Rk` at its status. Which is the headline? - writes the `claim` lines. A claim that cannot trace to a result is a named gap, a candidate to kick back to the Foundry.
4. **Equation + figure placement (atom).** Which display equation lands in this section, and which claim does it serve? Which already-built figure goes here? Equation-placement is a first-class decision made *now*, not deferred to drafting.
5. **Literature placement.** Walk the references on hand. For each, which section is its home (intro / related / discussion), and which claim does it support, contest, or contextualise? - writes the `cite` lines.

The ladder is re-enterable. A later answer may reopen an earlier one (a claim reveals the thesis was too narrow) - that is propagation, not failure. Lock a section node only when every claim traces to a result at a known status.

## The live preview loop (understand -> insert -> reconcile)
After each answer the author should *see* the paper assemble and be able to steer it:
- **understand** - run the backing (`project_blueprint.py`) to project `blueprint.md` into the graph, and view it (the viewer drills macro sections -> meso claims/figures -> atomic claim<->equation/evidence, and shows each claim's result + status).
- **insert** - the author answers the next question, or comments on a node in the viewer. Both write back to the blueprint.
- **reconcile** - a locked node is a committed element. Hand it to the [`reconciler`](canon://skills/reconciler.md), which propagates the change (lateral edits to the identity ledger, structural follow-ons queued) and rewrites the handoff last.

## The drafting gate
[`latex_forge`](canon://skills/latex_forge.md) drafts a section's prose **only from a locked blueprint node** whose every claim already traces to the ledger at a known status. Prose that asserts a claim absent from the node, or places an equation or figure the node did not, is a defect - caught against the blueprint, not discovered in review. When the network is locked, drafting is mechanical fill-in. That is the whole point.

## Invariants
- **Ask, do not assert.** Every identity, structure, claim, placement, and literature decision is the author's, surfaced as a concrete choice. You never decide the content.
- **Blueprint before prose.** No `.tex` section is drafted before its node locks.
- **Composition manifest, never prose.** A node holds commitments and intent. No inline duplication of a claim.
- **Core-outward.** Elicit the core first. Framing rings (conclusion, intro, abstract) are elicited last, because they summarise commitments that only exist once the inner nodes lock.
- **Provenance or it is a stub.** Every claim carries a `result:` and a status, or the node is not locked. Name the gap in the node.
- **One node per session.** The network grows piece by piece. Strays go to the backlog ([`focus_rail`](canon://skills/focus_rail.md)).
- **Reuse the figure.** Figures placed are the ones already built. Derive a variant from the original, never invent a parallel one.

## Backing
`scripts/project_blueprint.py` - projects `blueprint.md` into the [graph](canon://skills/canon_weaver.md) JSON the live viewer serves (`--out`), and prints a coverage report: section status counts, orphan claims (no result), sections without claims, premature framing (a core-outward guard), and the next core-outward node to grow. Run it after every answered decision so the picture is always current.

*Experimental and coarse by design. Refine in-session whenever the author's edits reveal a sharper question to ask, or a sharper standard for what a locked node must carry.*
