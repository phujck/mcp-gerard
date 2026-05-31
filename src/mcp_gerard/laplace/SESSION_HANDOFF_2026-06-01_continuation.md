# Continuation Handoff - 2026-06-01 (manuscript iteration + voice ingestion)

Follows `SESSION_HANDOFF_2026-05-31_overnight.md`. Read that first for the full
branch map. This records what changed overnight into 2026-06-01. Everything is
on the same isolated branches - `master`/`main` untouched.

## Headline

- The ANF manuscript is redrafted twice to your voice and **compiles to a
  12-page PDF**. It is a real first draft to hand-tune, not finished prose.
- A real voice corpus now grounds the drafting (brief v1 + v2, 55 sources).
- The engine backlog advanced (10 items across three batches, 76 tests green).

## Manuscript (orchestra branch draft/overnight, via worktree anf-draft-wt)

Work happens in the git worktree at `C:\Users\gerar\VScodeProjects\anf-draft-wt`
(draft/overnight checked out there, so it does not collide with the synthetics
branch on the main orchestra tree).

- Round 1 (`f1de205`): rebuilt all 11 sections + talk to the voice brief, killed
  the gross AI slop, cleared em-dashes and prose semicolons. Cost 1.17M tokens
  across 33 calls - retired that per-section design.
- Round 2 (`ca52dea`): one lean whole-manuscript pass (97k tokens, one agent)
  grounded in brief v2 - mechanism before name, register-collision, argument read
  as discovered, non-manufactured closure. The cadence to keep.
- Compile fixes (`612e472`): added `enumitem` (the A1-A7 list) and
  `amsthm`+`\newtheorem{theorem}` (the R6 classification). `main.tex` now builds.
  Remaining warnings are undefined `\cite` (provisional bib, no bibtex run) and
  cross-refs (resolve on a full `latexmk`). Open the PDF: `pdflatex main.tex` in
  `anf-draft-wt/adaptive-normal-form/manuscript/tex`.

**Still yours to do (unchanged):** the three sign-offs (N_0, "phases of
hierarchy" title, UEC), verify the borrowed 47-key bib, build Fig 4b. The draft
will need real hand-tuning - it is honest scaffolding in your register, not final.

## Voice corpus (engine branch dream/overnight)

- `VOICE_DRAFTING_BRIEF.md` (v1) and `VOICE_DRAFTING_BRIEF_v2.md` (`833867a`):
  the operational drafting brief. v2 adds THE INVARIANT VOICE ("one voice,
  varying amplitude - the manuscript register is the personal register with the
  scaffolding hidden"), register-transfer rules, 25 exemplars, expanded kill-list.
- 55 sources ingested across academic/personal/correspondence/admin. Distilled
  mechanics appended to the four register nodes with provenance handles.
  **Privacy held**: raw text and facts stayed in the gitignored cache, only
  distilled patterns entered canon.
- Cost lesson: the ingestion swarm read large PDFs whole and burned 6.6M tokens.
  Future ingestion must extract/trim text before the reader, not feed whole PDFs.

## Engine (dream/overnight)

- Backlog batches 1-3 committed (`209b4b9`, `ee6fb82`, plus the weave/firewall/
  canon_weaver commits): graph skill-ref + summary-mode fixes, orient generating
  rank, run_backing diagnostics, dreamer-surfaces-backlog, fitness-ignores-silence,
  README kernel-vs-routine, context_firewall sibling clause. 76 tests pass.
- The engine work on dream/overnight still merges into master as a clean
  fast-forward when you want it.

## Talk (draft/overnight)

- 20-slide Slidev skeleton at `adaptive-normal-form/talk/slides.md` + `spine.md`,
  voice-passed. Configured in `mcp-gerard/.claude/launch.json` as
  `anf-talk-slidev` (port 3030). Start with the preview tool. It is a SKELETON -
  the custom layouts and iframe asset paths need setup before it renders cleanly.
  Treat it as content to work from, not a finished deck.

## Not done / deferred

- Synthetics distillation did NOT complete (the evidence base is empty; the run
  did not survive an earlier pause). Branch `synthetics/distillation-2026-05-31`
  exists; the workflow script is saved under the session workflows dir. Re-run
  when wanted - it is cheap Haiku but not re-launched, to respect quota.

## Worktree note

`git worktree list` shows the main orchestra tree (synthetics branch) and the
`anf-draft-wt` worktree (draft/overnight). To remove the worktree later:
`git worktree remove anf-draft-wt`. The manuscript commits live on draft/overnight
regardless.

## Morning steps

1. `pdflatex main.tex` in the worktree, read the 12-page PDF, sign the three calls.
2. Hand-tune the prose - it is in your register but wants your hand.
3. Decide the bib (verify or replace the 47 provisional keys), build Fig 4b.
4. Start the talk deck, fix layouts/assets, iterate.
5. Merge dream/overnight into engine master if you accept the engine work.
