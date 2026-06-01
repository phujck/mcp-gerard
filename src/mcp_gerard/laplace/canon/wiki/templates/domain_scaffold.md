# Domain Scaffold (the extension protocol)

How the engine grows a new capability. The answer is not "add a subsystem" - it is **add a domain**. Manuscript drafting, talks, the website, a personal-knowledge store: each is a domain, scaffolded the same way and wired into the same relevance loop. This is scale-invariance applied to knowledge - the engine extends by self-similar accretion, not by special cases. [`project_scaffold`](canon://templates/project_scaffold.md) is the per-project sub-step inside a domain; this is the domain itself.

## The core tension this is built around
Two demands pull against each other, and the architecture already resolves them:

- **Keep context to the task.** `laplace_orient` returns only the task slice - the global foundation, the *one* active domain's axioms and project state, and a budgeted set of candidate skills. A domain's nodes are `scope: domain` and surface only when the goal matches. The whole canon is never loaded at once. This is the context firewall at the retrieval layer.
- **Make everything sensibly accessible.** `laplace_search` spans all canon, the [graph](canon://skills/canon_weaver.md) keeps it one connected object, and `canon_weaver` keeps the topology healthy. Nothing is hidden - it is just not *imposed*. Any node is one query away.

The same orient->slice mechanism works whether the canon is one project or an everything-machine. Adding a domain does not dilute the task context, because relevance ranking only ever surfaces the active slice. That is what makes unbounded growth safe.

## To add a capability domain
1. **Axioms.** Create `domains/<domain>/axioms.md` - the domain's strict rules, governing objects, and what it does *not* cover. Consult [`context_firewall`](canon://skills/context_firewall.md) first: prior material from another project enters only by re-derivation, and names are earned once the development produces what they denote.
2. **Projects.** Create `domains/<domain>/projects/<project>.md` per [`project_scaffold`](canon://templates/project_scaffold.md) - durable facts only, delegating live per-result status to the project's `HANDOFF.md`.
3. **Skills, forge-by-doing.** A domain's generating/staging/evaluating skills are earned from the friction of doing the act the first time, never imagined up front. Forge them as the work is done (the [core-outward trunk](canon://workflow/core_outward_trunk.md) governs the order).
4. **Register for relevance.** Add the domain to `index.yaml` under `domains:` with its tags, and tag its wiki nodes `scope: domain`. Now `orient` infers and surfaces the domain when a goal matches its language, and leaves it dormant otherwise.

## Dormant and future domains
- **web** is already a scaffolded-but-dormant domain (`domains/web/axioms`, the `phujck_github_io` project, three deprecated stage skills awaiting real backings). Turning website management on is *filling it in*, not a rewrite.
- A **personal-knowledge** domain (the everything-machine's memory of the author) is scaffolded the same way, reusing the [voice corpus](canon://voice_corpus/index.md) privacy/provenance pattern: cache raw source outside git, surface by relevance, never dump it into task context. The orient firewall is exactly what keeps a large personal store from leaking into an unrelated task.

## The scaling frontier (named, not yet built)
As the canon grows toward an everything-machine, three things must stay sharp - tracked in `AUTONOMY_BACKLOG.md`: bound `available_refs` and tighten relevance ranking so orient stays surgical; add **multi-domain composition** so a task spanning two domains (publish this result to the site) composes both slices rather than inferring one; and weight the personal store by relevance so accessibility never costs task focus.
