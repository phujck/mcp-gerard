"""Canon graph topology audit.

Projects the canon as a typed graph, identifies orphaned and disconnected
nodes, and for each orphan scans its source prose for canon entity names that
appear unlinked - candidate edges the weaver should add.

This is an audit script, not an auto-editor. It proposes; the operator decides.

Usage
-----
Run from the repo root with the packaged venv::

    ./.venv/Scripts/python.exe canon/skills/canon_weaver/scripts/weave_canon.py

Or as a module::

    python -m mcp_gerard.laplace.canon.skills.canon_weaver.scripts.weave_canon

Output
------
One health line, the orphan list, then for each orphan a list of OTHER canon
node labels/names that appear verbatim in its prose but are not yet linked.
These are the candidate edges. Verify each in the source file before adding.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Bootstrap the package import path when run as a plain script.
# The repo layout is src/mcp_gerard/..., so we walk up to the src/ dir.
# ---------------------------------------------------------------------------
_HERE = Path(__file__).resolve()
_SRC = _HERE.parents[6]  # scripts(0) -> canon_weaver(1) -> skills(2) -> canon(3) -> laplace(4) -> mcp_gerard(5) -> src(6)
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from mcp_gerard.laplace.graph import CanonGraph  # noqa: E402
from mcp_gerard.laplace.canon import get_canon, _split_frontmatter  # noqa: E402


def _prose_text(path: Path) -> str:
    """Return the body of a canon markdown file, stripped of frontmatter."""
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError):
        return ""
    _, body = _split_frontmatter(text)
    return body


def _existing_links(body: str) -> set[str]:
    """Collect all [[token]] and canon:// tokens already present in a prose body."""
    wikilink_re = re.compile(r"\[\[([^\]|]+)")
    canon_re = re.compile(r"canon://([A-Za-z0-9_./-]+)")
    linked: set[str] = set()
    for m in wikilink_re.finditer(body):
        token = m.group(1).strip().rsplit("/", 1)[-1].removesuffix(".md")
        linked.add(token.lower())
    for m in canon_re.finditer(body):
        raw = m.group(1).strip().rstrip(".").removesuffix(".md")
        token = raw.split("/", 1)[-1] if raw.startswith("skills/") else raw
        linked.add(token.lower())
    return linked


def _candidate_edges(orphan_id: str, orphan_path: Path, all_labels: dict[str, str]) -> list[str]:
    """Return canon labels/names that appear verbatim in the orphan's prose but are not yet linked."""
    body = _prose_text(orphan_path)
    if not body:
        return []
    already_linked = _existing_links(body)
    candidates: list[str] = []
    body_lower = body.lower()
    for label, node_id in all_labels.items():
        if node_id == orphan_id:
            continue
        label_lower = label.lower()
        if label_lower in already_linked:
            continue
        # Whole-word match: the label must appear as a distinct token in prose.
        pattern = r"(?<![A-Za-z0-9_])" + re.escape(label_lower) + r"(?![A-Za-z0-9_])"
        if re.search(pattern, body_lower):
            candidates.append(f"{label}  ({node_id})")
    return candidates


def audit() -> None:
    canon = get_canon()
    g = CanonGraph.from_canon(canon)
    h = g.health()

    # One-line health summary.
    orphan_count = len(h["orphans"])
    component_count = h["components"]
    passed = h["pass"]
    status = "PASS" if passed else "FAIL"
    print(
        f"health: {status} | components={component_count} | orphans={orphan_count} | "
        f"dangling={len(h['dangling'])} | dead_wood_linked={len(h['dead_wood_linked'])}"
    )

    if not h["orphans"]:
        print("\nNo orphans found.")
        return

    # Build a lookup: display label / skill name -> node id, for all real nodes.
    all_labels: dict[str, str] = {}
    for nid, node in g.nodes.items():
        if node.kind == "missing":
            continue
        all_labels[node.label] = nid
        # For skills: also index by their short name (id tail after "skill:").
        if node.kind == "skill":
            all_labels[nid.split(":", 1)[1]] = nid

    # Source path per node: wiki pages have a .path; skills have a SKILL.md.
    def _source_path(nid: str) -> Path | None:
        if nid.startswith("wiki:"):
            ref = nid[5:]
            node = canon.wiki.get(ref)
            return node.path if node else None
        if nid.startswith("skill:"):
            name = nid[6:]
            sk = canon.skills.get(name)
            return sk.skill_md if sk and sk.skill_md.exists() else None
        return None

    print(f"\nOrphaned nodes ({orphan_count}):")
    for oid in h["orphans"]:
        node = g.nodes[oid]
        print(f"  {oid}  [{node.kind}]  \"{node.label}\"")
        src = _source_path(oid)
        if src is None:
            print("    (no source file - cannot scan for candidates)")
            continue
        candidates = _candidate_edges(oid, src, all_labels)
        if candidates:
            print(f"    Candidate edges ({len(candidates)}):")
            for c in candidates:
                print(f"      + {c}")
        else:
            print("    No candidate edges found in prose.")


if __name__ == "__main__":
    audit()
