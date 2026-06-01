#!/usr/bin/env python3
"""Project a manuscript blueprint into the canon graph + a coverage report.

The backing for the ``manuscript_spine`` skill. It reads ``blueprint.md`` - the
plan elicited from the author, one section node at a time - projects it through
``graph.from_blueprint`` into the same node/edge JSON the live viewer and every
renderer already consume, and prints a coverage report: what is settled, stubbed,
or missing, and which core-outward node to grow next.

Usage:
    python project_blueprint.py <blueprint.md> [--out graph.json] [--format json|mermaid|canvas]

``--out`` writes the graph artifact (default JSON) for the viewer to serve.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from mcp_gerard.laplace import graph as _graph

# Framing rings are drafted last (core-outward); a framing section ahead of a
# still-stub core section is premature framing - the headline drift this guards.
_DONE = {"blueprinted", "drafted"}


def coverage(g: "_graph.CanonGraph") -> dict:
    """A blueprint coverage report derived from the projected graph."""
    sections = [n for n in g.nodes.values() if n.kind == "section"]
    claims = [n for n in g.nodes.values() if n.kind == "claim"]

    status_counts: dict[str, int] = {}
    for n in sections:
        st = n.status or "stub"
        status_counts[st] = status_counts.get(st, 0) + 1

    # which sections contain at least one claim / a headline claim
    claims_by_section: dict[str, list] = {}
    for e in g.edges:
        if e.rel == "contains" and g.nodes.get(e.dst) and g.nodes[e.dst].kind == "claim":
            claims_by_section.setdefault(e.src, []).append(g.nodes[e.dst])

    sections_without_claims = [s.label for s in sections if not claims_by_section.get(s.id)]
    headline_missing = [
        s.label for s in sections
        if claims_by_section.get(s.id) and not any(c.meta.get("headline") for c in claims_by_section[s.id])
    ]
    orphan_claims = [c.meta.get("text", c.label) for c in claims if not c.meta.get("result")]

    core_stub = [s for s in sections if (s.meta.get("ring") == "core") and (s.status or "stub") == "stub"]
    premature_framing = [
        s.label for s in sections
        if s.meta.get("ring") == "framing" and (s.status or "stub") in _DONE and core_stub
    ]

    # the next core-outward node to grow: the earliest section that is still a
    # stub, innermost ring first (core -> inner -> framing).
    ring_rank = {"core": 0, "inner": 1, "section": 2, "framing": 3}
    stubs = sorted(
        (s for s in sections if (s.status or "stub") == "stub"),
        key=lambda s: (ring_rank.get(s.meta.get("ring") or "section", 2), s.meta.get("order", 0)),
    )
    next_node = stubs[0].label if stubs else None

    return {
        "sections": len(sections),
        "claims": len(claims),
        "status_counts": status_counts,
        "sections_without_claims": sections_without_claims,
        "headline_missing": headline_missing,
        "orphan_claims_no_result": orphan_claims,
        "premature_framing": premature_framing,
        "next_to_grow": next_node,
        "complete": not stubs and not orphan_claims and not sections_without_claims,
    }


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Project a manuscript blueprint into the canon graph.")
    p.add_argument("blueprint", help="path to blueprint.md")
    p.add_argument("--format", default="json", choices=["json", "mermaid", "canvas", "obsidian"])
    p.add_argument("--out", default=None, help="write the graph artifact to a file (for the viewer)")
    args = p.parse_args(argv)

    bp = Path(args.blueprint)
    if not bp.exists():
        print(json.dumps({"error": f"blueprint not found: {bp}"}, indent=2))
        return 2

    res = _graph.render(args.format, blueprint=str(bp))
    if "error" in res:
        print(json.dumps(res, indent=2))
        return 2
    g = _graph.CanonGraph.from_blueprint(str(bp))
    cov = coverage(g)

    if args.out:
        artifact = res["artifact"]
        text = artifact if isinstance(artifact, str) else json.dumps(artifact, indent=2)
        Path(args.out).write_text(text, encoding="utf-8")

    print(json.dumps({"coverage": cov, "health": res.get("health"), "stats": res.get("stats")}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
