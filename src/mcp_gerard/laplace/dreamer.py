"""The Dreamer: the slow, self-refining loop.

Between sessions, the dreamer reads the telemetry-derived fitness assessment and
silently reparametrises the canon:

  * Deterministic curation (always safe, autonomous): apply the evidence-gated
    lifecycle transitions from ``assess`` - promote proven experimental skills to
    core, deprecate the unused or degraded - by writing the machine-owned
    ``lifecycle.yaml`` and committing it to the canon git history.

  * Generative forging (host-executed): the engine is self-contained and makes no
    external LLM API call. It hands a forging brief to the host - the most powerful
    model already in the loop, running locally - which drafts the SKILL.md and
    persists it via ``persist_forged_skill``, born ``experimental`` under the
    Probationary Protocol. Judgement routes to the host; determinism stays here.

Every mutation is a single scoped git commit, so ``rollback`` can undo it.
"""

from __future__ import annotations

import os
import re
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import yaml

from mcp_gerard.laplace import assess as _assess
from mcp_gerard.laplace import gitio
from mcp_gerard.laplace import telemetry as _telemetry
from mcp_gerard.laplace.canon import Canon, get_canon
from mcp_gerard.laplace.locking import CanonBusy, canon_lock

# ---------------------------------------------------------------------------
# backlog reader
# ---------------------------------------------------------------------------

# The backlog file lives alongside this module in the engine directory.
_BACKLOG_PATH = Path(__file__).resolve().parent / "AUTONOMY_BACKLOG.md"
_BACKLOG_HEADING_RE = re.compile(r"^##\s+(.+)", re.MULTILINE)


def _read_backlog() -> dict[str, Any]:
    """Return a digest of AUTONOMY_BACKLOG.md: its section headings and open items.

    Open items are lines that start with ``- **`` and are NOT struck-through
    (i.e. do not start with ``- ~~``).  The result is surfaced in the dream
    bundle under ``"backlog"`` so deferred items resurface each cycle without
    the dreamer having to remember them.
    """
    try:
        text = _BACKLOG_PATH.read_text(encoding="utf-8")
    except OSError:
        return {"available": False, "reason": "backlog file not readable"}

    sections = _BACKLOG_HEADING_RE.findall(text)
    open_items: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("- **") and not stripped.startswith("- ~~"):
            # Extract the bold label: text between the first pair of **
            m = re.match(r"- \*\*(.+?)\*\*", stripped)
            if m:
                open_items.append(m.group(1))
    return {
        "available": True,
        "sections": sections,
        "open_items": open_items,
        "path": str(_BACKLOG_PATH),
    }

# ---------------------------------------------------------------------------
# git helpers, scoped to the canon subtree
# ---------------------------------------------------------------------------


def _commit(root: Path, paths: list[Path], message: str) -> str | None:
    """Stage the given paths and commit via pure-Python git (Dulwich).

    Returns the new commit sha, or None if staging produced nothing. No
    subprocess, so no git.exe and no fsmonitor-daemon pipe to deadlock on.
    """
    return gitio.commit(root, paths, message)


# ---------------------------------------------------------------------------
# lifecycle persistence
# ---------------------------------------------------------------------------


def _lifecycle_path(canon: Canon) -> Path:
    return canon.root / "lifecycle.yaml"


def _save_lifecycle(canon: Canon, data: dict[str, Any]) -> Path:
    path = _lifecycle_path(canon)
    header = (
        "# Machine-owned lifecycle overlay. The dreamer writes this; do not hand-edit.\n"
        "# It overrides skill `status` in index.yaml based on measured fitness.\n"
    )
    # Atomic: write a uniquely-named temp in the same directory then os.replace,
    # so a lock-free reader (the canon loader) never sees a half-written,
    # invalid-YAML overlay. The temp name is unique (mkstemp) so the write is
    # self-contained and safe even off the lock. os.replace is atomic, but on
    # Windows it raises PermissionError if a reader holds the destination open
    # for the instant of the swap (CPython opens without share-delete), so retry
    # briefly - the reader's window is tiny.
    body = header + yaml.safe_dump(data, sort_keys=True)
    fd, tmpname = tempfile.mkstemp(dir=str(path.parent), prefix=path.name + ".", suffix=".tmp")
    tmp = Path(tmpname)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(body)
        for attempt in range(100):
            try:
                os.replace(tmp, path)
                return path
            except PermissionError:
                if attempt == 99:
                    raise
                time.sleep(0.01)
    finally:
        if tmp.exists():
            tmp.unlink()
    return path


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


# ---------------------------------------------------------------------------
# deterministic curation
# ---------------------------------------------------------------------------


def apply_transitions(canon: Canon, report: dict[str, Any]) -> dict[str, Any]:
    """Write recommended lifecycle transitions into lifecycle.yaml."""
    life = dict(canon.lifecycle or {})
    life.setdefault("skills", {})
    fitness_by = {r["name"]: r["fitness"] for r in report["skills"]}

    for t in report["transitions"]:
        name = t["name"]
        entry = life["skills"].get(name, {}) or {}
        entry["status"] = t["to"]
        entry["fitness"] = fitness_by.get(name)
        entry["updated"] = _now()
        entry.setdefault("history", []).append(
            {"from": t["from"], "to": t["to"], "reason": t["reason"], "at": _now()}
        )
        life["skills"][name] = entry

    path = _save_lifecycle(canon, life)
    return {"lifecycle_path": path, "applied": report["transitions"]}


# ---------------------------------------------------------------------------
# generative forging (optional)
# ---------------------------------------------------------------------------

_NAME_RE = re.compile(r"name:\s*([a-z0-9_]+)", re.IGNORECASE)


def _friction_from_report(report: dict[str, Any]) -> str:
    """Synthesise a friction brief from the fitness assessment when none is supplied.

    Lets the dreamer run autonomously: with no explicit friction it still hands the
    host a substantive, assessment-driven agenda.
    """
    lines: list[str] = []
    for r in report.get("refine_recommended", []):
        lines.append(f"- refine '{r['name']}': {r['refine_reason']} (signal {r['refine_signal']}).")
    unused = report.get("unused", [])
    if unused:
        lines.append("- long-unused skills to review for deprecation or merge: " + ", ".join(unused[:10]) + ".")
    if not lines:
        return "No explicit friction supplied. Review the session transcript for repeated manual work and conversational bottlenecks, and remedy them."
    return "Autonomous, assessment-driven refinement agenda:\n" + "\n".join(lines)


def forge_skill(
    canon: Canon,
    friction: str,
    transcript: str = "",
    model: str = "claude",
) -> dict[str, Any]:
    """Produce the forging brief for the host to execute directly and locally.

    The Laplace engine stands as its own object and makes no external LLM API call.
    Forging needs maximum thought and capacity, so the brief is handed to the host -
    the most powerful model already in the loop - which drafts the SKILL.md and
    persists it via ``persist_forged_skill``. ``model`` is an advisory hint only.
    """
    persona = canon.resolve("agents/the_dreamer.yaml")[1] if (canon.root / "agents" / "the_dreamer.yaml").exists() else ""
    instructions = (
        "You are the host - the most powerful model already in the loop, running "
        "locally. Execute this brief directly and with maximum rigour. Do not "
        "delegate it to an external API.\n\n"
        "Forge ONE new skill (or refine the named weak one) that would permanently "
        "eliminate the described friction. Write a complete SKILL.md with YAML "
        "frontmatter containing `name` (snake_case) and `description`, then a concise "
        "protocol. The skill is born EXPERIMENTAL. Persist it via "
        "persist_forged_skill(canon, content) and register its activity and tags in "
        "index.yaml. Then commit the canon (a single revertible commit).\n\n"
        f"OBSERVED FRICTION:\n{friction}\n\n"
        f"RECENT TRANSCRIPT (excerpt):\n{transcript[:4000]}"
    )
    return {
        "forged": False,
        "mode": "host_forge",
        "reason": "Laplace engine is self-contained; the host model forges directly and locally.",
        "model_hint": model,
        "persona": persona,
        "instructions": instructions,
    }


def persist_forged_skill(canon: Canon, content: str) -> dict[str, Any]:
    """Persist a host-drafted SKILL.md and register it experimental in lifecycle.

    Called by the host after it has drafted the skill from the forge brief.
    """
    m = _NAME_RE.search(content)
    if not m:
        return {"forged": False, "reason": "no name found in draft frontmatter"}
    name = m.group(1).lower()
    try:
        with canon_lock():
            # Re-read the canon under the lock so the lifecycle mutation is built
            # on current state, not a snapshot another agent may have superseded.
            fresh = get_canon(fresh=True)
            skill_dir = fresh.root / "skills" / name
            skill_dir.mkdir(parents=True, exist_ok=True)
            skill_md = skill_dir / "SKILL.md"
            skill_md.write_text(content, encoding="utf-8")

            # Register as experimental in the machine-owned lifecycle overlay.
            life = dict(fresh.lifecycle or {})
            life.setdefault("skills", {})
            life["skills"][name] = {
                "status": "experimental",
                "forged": _now(),
                "history": [{"from": None, "to": "experimental", "reason": "forged by dreamer (host-executed)", "at": _now()}],
            }
            _save_lifecycle(fresh, life)
            return {"forged": True, "name": name, "path": skill_md, "skill_md": skill_md}
    except CanonBusy as exc:
        return {"forged": False, "busy": True, "reason": str(exc)}


# ---------------------------------------------------------------------------
# the cycle
# ---------------------------------------------------------------------------



def dream(
    apply: bool = True,
    forge: bool = False,
    friction: str = "",
    transcript: str = "",
    model: str = "claude",
    commit: bool = True,
) -> dict[str, Any]:
    """Run the R&R cycle. Deterministic curation by default; forging on request.

    The read-assess-mutate-commit critical section runs under the cross-process
    canon lock, so concurrent dreams from multiple agents serialise instead of
    clobbering lifecycle.yaml. On contention it returns a structured busy result
    rather than blocking.
    """
    since = _telemetry.last_dream_ts()  # only assess events since the last dream
    try:
        with canon_lock():
            return _dream_locked(apply, forge, friction, transcript, model, commit, since)
    except CanonBusy as exc:
        return {"busy": True, "error": str(exc), "boundary_advanced": False}


def _dream_locked(
    apply: bool,
    forge: bool,
    friction: str,
    transcript: str,
    model: str,
    commit: bool,
    since: str | None,
) -> dict[str, Any]:
    """The dream body, run while holding the canon write lock."""
    canon = get_canon(fresh=True)
    report = _assess.assess(canon, since=since)
    out: dict[str, Any] = {
        "assessment": {
            "transitions": report["transitions"],
            "refine_recommended": report.get("refine_recommended", []),
            "unused": report["unused"],
            "events_seen": report["events_seen"],
        },
        "backlog": _read_backlog(),
    }
    changed: list[Path] = []

    if apply and report["transitions"]:
        res = apply_transitions(canon, report)
        out["curation"] = {"applied": res["applied"]}
        changed.append(res["lifecycle_path"])

    if forge:
        canon = get_canon(fresh=True)  # pick up lifecycle write
        brief_friction = friction or _friction_from_report(report)
        # Self-contained engine: hand the brief to the host model to execute
        # directly and locally. The host persists via persist_forged_skill and
        # commits separately, so nothing is auto-committed here.
        out["forge"] = forge_skill(canon, brief_friction, transcript, model)

    if commit and changed:
        msg = _audit_message(out)
        sha = _commit(canon.root, changed, msg)
        out["commit"] = sha

    get_canon(fresh=True)  # ensure subsequent reads see the new canon

    # Window-leak guard: advance the assessment boundary only when the dream
    # actually mutated canon - i.e. applied a lifecycle transition. A preview
    # (apply=False) or a dream that found nothing actionable must NOT consume the
    # window. Its evidence is unspent and has to stay visible to the next real
    # dream, or genuine friction slips behind the boundary, assessed once and never
    # acted on (AUTONOMY_BACKLOG, "No-op dream orphans evidence"). The earlier
    # event-count floor missed the common case: an apply=False preview over a full
    # window stamped the boundary and orphaned every event in it. Forging is
    # host-executed and committed separately, so it does not advance here either.
    did_apply = bool(apply and report["transitions"])
    if did_apply:
        _telemetry.log("dream_complete", events_seen=report["events_seen"], since=since)
        out["boundary_advanced"] = True
    else:
        out["boundary_advanced"] = False
    return out


def _audit_message(out: dict[str, Any]) -> str:
    parts = ["dream:"]
    for t in out.get("curation", {}).get("applied", []):
        parts.append(f"{t['name']} {t['from']}->{t['to']}")
    f = out.get("forge", {})
    if f.get("forged"):
        parts.append(f"forge {f['name']}")
    return " ".join(parts) if len(parts) > 1 else "dream: no-op"


def rollback(ref: str) -> dict[str, Any]:
    """Revert a previous dreamer commit (scoped to canon files), via Dulwich."""
    try:
        with canon_lock():
            canon = get_canon()
            ok, detail = gitio.revert(canon.root, ref)
            get_canon(fresh=True)
            return {"ref": ref, "ok": ok, "detail": detail}
    except CanonBusy as exc:
        return {"ref": ref, "ok": False, "busy": True, "error": str(exc)}
