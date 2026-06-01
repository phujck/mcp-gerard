"""The manuscript blueprint projection: blueprint.md -> the canon graph.

The blueprint is the plan that exists before any .tex. Projecting it through the
same node/edge model as a compiled manuscript is what makes the paper navigable
while it is still being elicited - one model, three sources (canon, manuscript,
blueprint).
"""

from __future__ import annotations

import importlib.util

from mcp_gerard.laplace import canon as canonmod
from mcp_gerard.laplace import graph

SAMPLE = r"""# Blueprint: Adaptive Normal Form

- thesis: One eigenvalue prices both walls of the viable window.
- frame: The UV catastrophe was a sum over modes.

## sec:intro | Introduction | ring:framing | status:stub
intent: open on the UV-catastrophe analogy, then hand to the model
- claim: A locally reasonable rule is lethal in aggregate | result:R0 | status:open-gap
- cite: Rayleigh1900 | relation:context

## sec:model | The Adaptive Normal Form | ring:core | status:blueprinted
intent: name the loop and its objects, derive the return map
- claim: The loop collapses to one scalar return map | result:R1 | status:established | headline
- equation: eq:returnmap | e_{k+1} = (r - K_N) e_k - alpha c_N e_k^3 | serves:1
- figure: fig:1 | the adaptive loop
- claim: psi classifies the phase | result:R2 | status:established
- equation: eq:chi | \chi_N \sim N^\psi | serves:psi classifies
"""


def _write(tmp_path):
    bp = tmp_path / "blueprint.md"
    bp.write_text(SAMPLE, encoding="utf-8")
    return bp


def test_from_blueprint_builds_manuscript_graph(tmp_path):
    g = graph.CanonGraph.from_blueprint(str(_write(tmp_path)))
    kinds: dict[str, int] = {}
    for n in g.nodes.values():
        kinds[n.kind] = kinds.get(n.kind, 0) + 1
    assert kinds.get("section") == 2
    assert kinds.get("claim") == 3
    assert kinds.get("equation") == 2
    assert kinds.get("figure") == 1
    assert kinds.get("citation") == 1

    rels = {e.rel for e in g.edges}
    assert {"precedes", "contains", "supported_by", "cites"} <= rels

    headline = [n for n in g.nodes.values() if n.kind == "claim" and n.meta.get("headline")]
    assert headline and headline[0].meta["result"] == "R1"
    assert headline[0].meta["status"] == "established"


def test_render_blueprint_json_and_mermaid(tmp_path):
    bp = str(_write(tmp_path))
    j = graph.render("json", blueprint=bp)
    assert "artifact" in j and j["artifact"]["nodes"]
    m = graph.render("mermaid", blueprint=bp)
    assert isinstance(m["artifact"], str) and m["artifact"].startswith("graph")


def test_blueprint_coverage_report(tmp_path):
    bp = str(_write(tmp_path))
    script = canonmod._PACKAGED_CANON / "skills" / "manuscript_spine" / "scripts" / "project_blueprint.py"
    spec = importlib.util.spec_from_file_location("project_blueprint", script)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    g = graph.CanonGraph.from_blueprint(bp)
    cov = mod.coverage(g)
    assert cov["sections"] == 2
    assert cov["claims"] == 3
    assert cov["next_to_grow"] == "Introduction"  # the only stub
    assert cov["complete"] is False
    # the model section has a headline claim; the intro does not yet
    assert "Introduction" in cov["headline_missing"]
