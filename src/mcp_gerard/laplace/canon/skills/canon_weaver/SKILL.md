---
name: canon_weaver
description: Audits the canon graph topology and proposes prose-supported links to repair orphaned and disconnected nodes.
---

# Canon Weaver

The canon is a typed graph. Nodes are wiki pages, skills, domains, and agent personas. Edges are the [[name]] wikilinks and canon:// references woven into their prose. `graph.py` measures topology health - orphans, dangling links, disconnected components - but nothing repairs it. The canon-weave task has had no owner. This skill is that owner.

[[global_weaver]] handles the manuscript orchestra: cross-references and labels across `.tex` papers via `scripts/weave_orchestra.py`. Canon Weaver is its counterpart for the canon itself.

## Usage

```
laplace_run(skill="canon_weaver", target="canon")
```

This executes `scripts/weave_canon.py` and prints an audit report: the health line, the orphan list, and for each orphan the candidate edges it found - canon node names that appear in the orphan's prose but are not yet linked.

For the topology summary alone:

```
laplace_graph(format="json")
```

and read the `health` key directly.

## Protocol

1. Run the audit script and read its output.
2. For each candidate edge, open the orphan's source file and verify the named entity genuinely appears in running prose - not in a code fence or heading fragment.
3. Add a `[[skill_name]]` wikilink or a `canon://ref` where the prose already names the target. Never invent a link the prose does not support.
4. **Hard firewall**: never link project lore (domain-scoped wiki pages, project notes) into a global or core node. The weave direction is outward - a domain page may reference a global concept, not the reverse.
5. Re-run the audit. Report the component-count before and after.

Address all genuine orphans before closing an engine self-development session.
