---
name: session_closer
description: "[EXPERIMENTAL] Close a session in two re-entrant tiers: a cheap idempotent CHECKPOINT (log skills, reconcile, rewrite the handoff - safe to re-run after trailing messages, processing only new evidence) and an infrequent evidence-gated STRUCTURAL DREAM (lifecycle transitions, forge, backlog work over a cumulative window), so session evidence compounds into canon fitness without a premature close orphaning the messages that follow it."
---
# session_closer

**Status: [EXPERIMENTAL]**

> **Earned name: Sawbones.** The index keeps `session_closer` for discovery, but the protocol earned a bedside name the session it was forged. The close is the suture - run so the session does not bleed out: nothing dropped, nothing stale, no evidence left to evaporate on the floor.

## Purpose
A session ends the same way every time, and the close was reconstructed from first principles each time it happened. An ad hoc close drops steps - most damagingly per-skill logging, which starves the fitness loop at the source. This skill makes the close a single named protocol any model in the loop follows identically.

But a close is not a clean cliff edge: **the author reliably sends a few more messages after they think the session is done.** A one-shot close that assesses, dreams, and stamps the boundary will orphan that trailing evidence past the window - the exact bug that wrongly deprecated a structural skill on a thin slice. So the close is **two re-entrant tiers**: a cheap checkpoint that is safe to run repeatedly, and a heavy structural dream that is gated and infrequent.

This skill is the host-facing counterpart to the END-OF-SESSION ROUTINE in `canon://agents/the_dreamer.yaml`.

## Who runs it, and when
- **The session-closer only** - whoever is wrapping the session (the orchestrator/principal), never a result-worker mid-ladder.
- **Tier 1 every time the work is set down**, and again after any trailing messages - it is idempotent.
- **Tier 2 occasionally** - on an explicit "structurally close / dream this" intent, or when enough cumulative evidence has built since the last structural dream. Not on every checkpoint.

## Tier 1 - Checkpoint (cheap, idempotent, run freely)
Processes only the telemetry since the last checkpoint marker, so re-running it after trailing messages just folds in the new slice. It **never deprecates and never advances the structural-dream boundary.**
1. **Settle and commit.** Stage only the finished artifacts of the element being closed; one scoped commit, the project's existing convention.
2. **Log every skill used - honestly.** For each skill whose protocol was actually followed since the last checkpoint, call `laplace_log` with +1 if following it produced a verify-passing result, -1 if it caused friction. Tie each signal to the artifact it produced. Self-log behavioural +1s for verify-passing usage - the author under-vocalises praise, so silence must not read as failure (the "silence is not dissatisfaction" backstop, also enforced structurally in `assess`).
3. **Reconcile + rewrite the handoff last.** Run the [`reconciler`](canon://skills/reconciler.md) pass on the active project: classify what changed (lateral / downward / upward), apply lateral changes to the ledgers, queue structural follow-ons, and rewrite the active project's `HANDOFF.md` **last** (see [core-outward trunk](canon://workflow/core_outward_trunk.md)). The brief is the next session's starting context - the observer is inside the box.
4. **Sync external state.** Bring any out-of-canon record current - auto-memory, project index, dashboards, and the shared literature corpus (a [`corpus_librarian`](canon://skills/corpus_librarian.md) provenance-update for the citekeys this session drew on). The **canon project node** holds only durable facts and delegates per-result status to the `HANDOFF.md` - confirm it still points to the handoff rather than restating a specific active result.

## Tier 2 - Structural dream (heavy, evidence-gated, revertible)
The lifecycle + forge pass. Run it deliberately, not on every checkpoint.
1. **Assess over the cumulative window.** Call `laplace_assess`. Judge lifecycle on accumulated evidence, not a thin recency slice - a rail that turns over once per project-stage must be judged on a project-lifetime window, not a since-last-dream cut.
2. **Dream.** Call `laplace_dream`. Apply deterministic curation. A structural skill (one the global workflow nodes or a core skill reference) is never deprecated on silence - only on a genuine negative outcome (`assess` enforces this). Forge only against friction that genuinely recurred. Any forge brief is `host_forge` - execute it inline and locally, never via an external API.
3. **Work the backlog.** `AUTONOMY_BACKLOG.md` is the self-improvement queue. The dream surfaces its headers; when running with structural intent, take the highest-impact item or a flagged refine, do it, and record what was cleared. This is where evolutionary self-improvement actually happens - bounded, revertible, one commit each.

## Invariants
- **Tier 1 is idempotent and never deprecates.** Re-running it after trailing messages is safe and folds in only new evidence. It must not stamp the structural boundary.
- **Tier 2 is gated and cumulative.** It runs occasionally, judges on accumulated evidence, and never deprecates a structural skill on silence.
- **Logging precedes any dream.** A dream over an unlogged window dreams over nothing.
- **The closer closes.** A result-worker mid-ladder does not run this. Role boundary before sequence.
- **Forge briefs execute inline.** The engine is self-contained; the host drafts and persists. No external LLM call.
- **The handoff is rewritten last, every checkpoint.** If it was not touched, the checkpoint is incomplete.

## Backing
No deterministic script yet. A natural backing is a close-out verifier: given the session's tool-call log, confirm each used skill received a `laplace_log` event and the handoff was touched after the last result commit, emitting PASS/FAIL on checkpoint completeness. The judgement - which skills were genuinely used, what each earned, when to run Tier 2 - stays with the closer.
