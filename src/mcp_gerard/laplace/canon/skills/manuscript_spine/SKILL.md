---
name: manuscript_spine
description: "[EXPERIMENTAL] Own the Spine stage of the trunk — build and grow the manuscript blueprint (the NMD): an ordered set of section nodes, each pinning its claims (traced to a result ledger at status), its placed equations and figures, and its connective intent, projectable as a live preview network grown one node per session, so the final prose pass is mechanical fill-in."
---
# manuscript_spine

**Status: [EXPERIMENTAL]**

## Purpose
The [core-outward trunk](canon://workflow/core_outward_trunk.md) has three stages - Foundry, **Spine**, Trunk. The Foundry rails have owners ([`result_foundry`](canon://skills/result_foundry.md), [`literature_scout`](canon://skills/literature_scout.md), [`numerical_evidence`](canon://skills/numerical_evidence.md)). The Trunk has [`latex_forge`](canon://skills/latex_forge.md). The Spine stage had **no owner** - the narrative master doc (NMD) is specified in [Evidence-Schema Flow](canon://workflow/evidence_schema_flow.md) but nothing built it, grew it, or rendered it. This skill is that owner. It is the manuscript peer of the talks domain's `talk_spine_generator` ([Talk From Manuscript](canon://domains/talks/projects/talk_from_manuscript.md)).

The failure this prevents: drafting straight from the result ledgers into `.tex` prose, front-to-back from frozen inputs - exactly what the trunk forbids, and the observed way autonomous paper-building fails. The fix transfers what already works in the talk pipeline. A talk spine is a beat list - atomic units, each carrying one idea, an intent, and a provenance (a status traced to a result ledger and a figure label) - and slide-building became trivial because every unit pre-committed its claim, its figure, and its status before any prose existed. The manuscript blueprint is the same object: section nodes instead of beats, with equation-placement added as a first-class commitment, grown piece by piece across sessions.

## The blueprint is a composition manifest, not prose
The NMD holds **commitments and intent, never the sentences**. Result-local prose lives in the faces, section-level connective intent lives here, and the rendered draft is assembled later. Duplication is forbidden and checkable - a claim stated in a node may not also be drafted inline in the node (see [Evidence-Schema Flow](canon://workflow/evidence_schema_flow.md)). The blueprint lives at `workflow/state/blueprint.md` in the project root, beside the result ledgers it reads.

## What a section node carries (the five faces)
A section node is **locked** when it carries all five, every claim tracing to a ledger at a known status. Until then it is a `stub` and its gap is named, never smoothed.

1. **Claims** - the ordered statements the section lands, each traced by `result_id` to its R-ledger **at its status** (*established* / *partial-under-a-stated-restriction* / *open-gap* - a verification verdict, never an assertion). A claim with no `result_id` is a stub claim, not a locked one.
2. **Equations** - the display equations placed in this section, each tied to the claim it serves. *Equation-placement is a first-class decision made here* - which equation lands in which section - not deferred to drafting. Record the equation in its canonical form and the label it will carry.
3. **Figures** - the figure labels placed here, reusing the figures already built (never inventing parallel ones). The figure travels as itself - a plot is read, not distilled.
4. **Intent** - one line on what the section must land, and the transition to its neighbour. This is the connective-prose intent, the rhetorical move pre-committed, **never the prose**.
5. **Status + provenance** - the node status (`stub` / `blueprinted` / `drafted`) and a provenance flag (`SETTLED` when every claim is ledger-backed at a known status, `PROVISIONAL` when any claim is borrowed or unverified, exactly as the talk spine flags its beats).

## The pass (in order)
1. **Load the ledger bus, not the prose.** Read the result ledgers (R1..Rn four-face docs), the identity/glossary ledger, and the figure faces - plus the existing `blueprint.md` if one exists. **Never** load the prior manuscript prose or a superseded lineage (the [context firewall](canon://skills/context_firewall.md) holds).
2. **Pin the spine first, once.** If the global frame is not yet set, settle it before any node: the thesis sentence and the core-outward **section order**, drawn from the identity ledger and the one-paragraph statement of the result. The arc is the work - the nodes are its projection. This is the Spine stage proper.
3. **Grow one node this session, core-outward.** Build or sharpen a single section node, in core-outward order (core results first, framing rings last). Fill its five faces. Every claim must trace to a `result_id` at a known status - if it cannot, the node stays a `stub` and the gap is a first-class entry, a candidate to kick back to the Foundry.
4. **Place equations and figures as commitments.** Decide which display equation and which figure lands in this section, and record it in the node. A floating equation with no home, or a figure placed in no section, is a defect the projection will surface.
5. **Lock the node** only when every claim traces to the ledger at a known status. Mark its status and provenance. A node that merely reads well is not locked.
6. **Project the live preview network.** Run the backing to render the blueprint as the manuscript graph plus a coverage report, and eyeball it (the [graph projector](canon://skills/canon_weaver.md) idiom, manuscript side). Read the topology: orphan sections, claims with no result, equations with no home, figures unplaced, rings out of core-outward order. The growing graph *is* the live preview the author steers by.
7. **Hand off on sign-off.** A locked node is a committed element - hand it to the [`reconciler`](canon://skills/reconciler.md), which propagates and rewrites the brief. The next session grows the next node.

## The organisation face (separable from content)
The blueprint carries a **running-order table** (Section | what it lands | ring | status) at its head - the manuscript analogue of the talk spine's timing table. Organisation lives here, distinct from content. Reordering sections edits this table, never a node's five faces. A consolidated **open-gaps / PROVISIONAL ledger** sits at the foot, so every stub claim and borrowed move is visible in one place, not buried in a node.

## The drafting gate (how this makes prose trivial)
`latex_forge` drafts the prose for a section **only from a locked blueprint node** whose every claim already traces to the ledger at a known status. The blueprint is the input to the Trunk stage. Prose that asserts a claim absent from the node, or places an equation or figure the node did not, is a defect - caught against the blueprint, not discovered in review. When the whole network is locked, the prose pass is mechanical fill-in against pre-committed commitments. That is the point - the live preview network of the whole paper, built piece by piece, so drafting becomes trivial.

## Invariants
- **Blueprint before prose.** No `.tex` section is drafted before its node locks. This is the front-to-back firewall.
- **Composition manifest, never prose.** A node holds commitments and intent. Duplication of a claim into inline prose is forbidden.
- **Core-outward.** Grow from the settled core. The framing rings - conclusion, introduction, abstract - are blueprinted **last**, because they summarise commitments that only exist once the inner nodes lock.
- **Provenance or it is a stub.** Every claim carries a `result_id` and a status, or the node is not locked. Name the gap in the node, never in a footnote you will forget.
- **One node per session.** The network grows piece by piece. Strays go to the backlog ([`focus_rail`](canon://skills/focus_rail.md)).
- **Reuse the figure.** Figures placed in nodes are the ones already built. If a section needs a variant, derive it from the original, do not invent a parallel one.

## Backing
`scripts/project_blueprint.py` - a **forge-by-doing target**, lifted once it proves on the first real blueprint (the [`figure_standard`](canon://skills/figure_standard.md) / `anf_style.py` pattern). It parses `blueprint.md`'s structured section nodes and emits the [`laplace_graph`](canon://skills/canon_weaver.md) manuscript schema (sections / claims / equations / figures / citations) **before the `.tex` exists**, plus a coverage report: settled vs stubbed vs missing nodes, orphan claims, unplaced equations and figures, sections carrying no claim, and rings out of core-outward order. Until it is forged, keep `blueprint.md` structured so the projection is mechanical, and project by hand against `laplace_graph --manuscript` once a `.tex` skeleton exists. The semantic work - which claim, which equation, which ordering - stays judgement.

*Experimental and coarse by design. Refine in-session whenever the author's edits reveal a sharper standard for what a section node must carry.*
