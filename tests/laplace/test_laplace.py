"""Tests for the Laplace Engine: canon, verify, assessment, dreamer, render."""

from __future__ import annotations

import json
import shutil

import pytest
import yaml

from mcp_gerard.laplace import assess, dreamer, gitio, render, telemetry, verify
from mcp_gerard.laplace import canon as canonmod
from mcp_gerard.laplace.canon import Canon

pytestmark = pytest.mark.unit


SEED_TEX = r"""\documentclass{article}
\begin{document}
\section{Introduction}
We analyze the color of the system; this is crucial --- really.
\begin{equation}\label{eq:used} x = 1 \end{equation}
As shown in \ref{eq:used}, the result holds.
\begin{equation}\label{eq:orphan} y = 2 \end{equation}
We also cite \ref{eq:missing} here.
\end{document}
"""


# ---------------------------------------------------------------------------
# canon
# ---------------------------------------------------------------------------


def test_canon_loads_skills_and_wiki():
    c = Canon.load()
    assert len(c.skills) >= 10
    assert len(c.wiki) >= 8
    # activities and statuses are constrained vocabularies
    assert {s.activity for s in c.skills.values()} <= {"generating", "staging", "evaluating"}
    assert {s.status for s in c.skills.values()} <= {"experimental", "core", "deprecated"}


def test_canon_resolve_variants():
    c = Canon.load()
    # canon:// ref, bare wiki ref, and bare skill name all resolve
    assert "Laplace Voice" in c.resolve("canon://aesthetics/voice_and_style")[1]
    assert c.resolve("aesthetics/voice_and_style")[1]
    assert c.resolve("epistemic_ledger")[0].name == "SKILL.md"
    with pytest.raises(KeyError):
        c.resolve("canon://nope/nothing")


def test_get_canon_reloads_on_disk_change(tmp_path, monkeypatch):
    dst = tmp_path / "canon"
    shutil.copytree(canonmod._PACKAGED_CANON, dst)
    monkeypatch.setenv("LAPLACE_CANON", str(dst))
    first = canonmod.get_canon(fresh=True)
    # An unchanged canon returns the cached object - no needless reload.
    assert canonmod.get_canon() is first
    # A skill forged on disk (host-forge style) is visible without fresh=True:
    # the live tools pick up the change within the running process.
    probe = dst / "skills" / "probe_skill"
    probe.mkdir(parents=True)
    (probe / "SKILL.md").write_text(
        "---\nname: probe_skill\ndescription: probe\n---\n# Probe [EXPERIMENTAL]\n",
        encoding="utf-8",
    )
    reloaded = canonmod.get_canon()
    assert reloaded is not first
    assert "probe_skill" in reloaded.skills


def test_orient_infers_domain_and_ranks_skills():
    c = Canon.load()
    b = c.orient("lint the latex voice and check derivations for the phases of hierarchy")
    assert b["domain"] == "synthetics"
    # domain axioms + project are loaded with content
    refs = {d["ref"] for d in b["domain_context"]}
    assert any("synthetics/axioms" in r for r in refs)
    # skills are bucketed by activity (generating|staging|evaluating)
    names = {s["name"] for bucket in b["skills"].values() for s in bucket}
    assert {"latex_forge", "epistemic_ledger"} & names


def test_orient_surfaces_generating_skills_for_generation_goals():
    """orient must return a non-empty generating bucket for goals that signal
    creation intent, even when no generating-skill tag directly matches the
    goal tokens.  Non-generating goals must not pick up generating skills.
    """
    c = Canon.load()
    # Goals that clearly signal writing/creation intent.
    for goal in ("draft the introduction", "write a new section", "generate ideas"):
        b = c.orient(goal)
        gen = b["skills"]["generating"]
        assert gen, (
            f"orient returned empty generating bucket for goal {goal!r}; "
            f"expected at least one generating skill"
        )
        assert all(s["activity"] == "generating" for s in gen), (
            "non-generating skill leaked into generating bucket"
        )
    # A purely evaluative goal must NOT artificially inject generating skills.
    b_eval = c.orient("check the equation labels")
    # The evaluating bucket should be populated; generating may be empty.
    assert b_eval["skills"]["evaluating"], "evaluating bucket empty for an evaluative goal"


# ---------------------------------------------------------------------------
# verify (mirrors legacy ledgers)
# ---------------------------------------------------------------------------


def test_verify_flags_seeded_errors(tmp_path):
    tex = tmp_path / "seed.tex"
    tex.write_text(SEED_TEX, encoding="utf-8")
    rep = verify.verify(str(tex))
    assert rep["passed"] is False
    kinds = {f["kind"] for f in rep["checks"]["voice"]["violations"]}
    assert {"americanism", "ai_slop", "vonnegut"} <= kinds
    assert "eq:orphan" in rep["checks"]["epistemic"]["orphans"]
    assert any(b["ref"] == "eq:missing" for b in rep["checks"]["crossref"]["broken_refs"])


def test_epistemic_counts_eqref_not_just_ref(tmp_path):
    # A label referenced only via \eqref must not be reported as an orphan.
    tex = tmp_path / "eqref.tex"
    tex.write_text(
        r"\begin{equation}\label{eq:a} x=1 \end{equation} see \eqref{eq:a}." + "\n",
        encoding="utf-8",
    )
    rep = verify.verify(str(tex), checks=["epistemic"])
    assert rep["checks"]["epistemic"]["orphans"] == []


def test_voice_flags_unicode_emdash(tmp_path):
    tex = tmp_path / "uni.tex"
    tex.write_text("This clause — an aside — lingers.\n", encoding="utf-8")
    rep = verify.verify(str(tex), checks=["voice"])
    assert any(f["kind"] == "vonnegut" for f in rep["checks"]["voice"]["violations"])


def test_voice_ignores_markdown_syntax_noise(tmp_path):
    md = tmp_path / "syntax.md"
    md.write_text(
        "---\nname: probe\n---\n"
        "| A | B |\n| --- | --- |\n"
        "```text\ncmd --source-handle TEST --cache-root tmp\n```\n"
        "Clean prose stays here.\n",
        encoding="utf-8",
    )
    rep = verify.verify(str(md), checks=["voice"])
    assert rep["checks"]["voice"]["violations"] == []


def test_voice_ignores_markdown_horizontal_rule(tmp_path):
    md = tmp_path / "rule.md"
    md.write_text(
        "# Heading\n\nClean prose above.\n\n---\n\nClean prose below.\n",
        encoding="utf-8",
    )
    rep = verify.verify(str(md), checks=["voice"])
    assert rep["checks"]["voice"]["violations"] == []


def test_verify_clean_file_passes(tmp_path):
    tex = tmp_path / "clean.tex"
    tex.write_text(
        r"\section{Intro}\begin{equation}\label{eq:a} x=1 \end{equation}"
        r" see \ref{eq:a}." + "\n",
        encoding="utf-8",
    )
    rep = verify.verify(str(tex))
    assert rep["passed"] is True
    assert rep["issue_count"] == 0


# ---------------------------------------------------------------------------
# assessment + lifecycle
# ---------------------------------------------------------------------------


@pytest.fixture
def isolated_state(tmp_path, monkeypatch):
    monkeypatch.setenv("LAPLACE_STATE", str(tmp_path / "state"))
    telemetry.clear()
    yield
    telemetry.clear()


def test_fitness_promotes_and_deprecates(isolated_canon):
    # isolated_canon seeds both skills experimental, so good evidence promotes
    # latex_forge to core and sustained non-use deprecates css_forge.
    for _ in range(6):
        telemetry.log("verify_check", skill="latex_forge", check="voice", passed=True)
    for _ in range(9):
        telemetry.log("orient", domain="web", offered=["css_forge"])
    rep = assess.assess(Canon.load())
    moves = {t["name"]: t["to"] for t in rep["transitions"]}
    assert moves.get("latex_forge") == "core"
    assert moves.get("css_forge") == "deprecated"
    assert "css_forge" in rep["unused"]


def test_recently_committed_skill_is_spared_from_deprecation(isolated_canon):
    # A skill refined/forged in the assess window is spared the "offered, never used"
    # deprecation: refining a SKILL.md emits no usage event, so the commit is the
    # evidence of work the telemetry missed. A blind dream must not deprecate the
    # newest canon the same window it was built.
    dst = isolated_canon
    md = dst / "skills" / "css_forge" / "SKILL.md"
    md.write_text(md.read_text(encoding="utf-8") + "\n<!-- refined this window -->\n", encoding="utf-8")
    gitio.commit_all(dst, "refine css_forge")  # a non-root commit in-window

    # Both offered+unused, both non-structural. css_forge was just committed;
    # html_mechanic was not. (latex_forge is unusable here as a deprecation
    # example: it is the structural Trunk owner, so it is protected regardless.)
    for _ in range(9):
        telemetry.log("orient", domain="web", offered=["css_forge"])
        telemetry.log("orient", domain="web", offered=["html_mechanic"])
    rep = assess.assess(Canon.load())
    moves = {t["name"]: t["to"] for t in rep["transitions"]}
    assert moves.get("css_forge") != "deprecated"        # spared by the commit grace
    assert moves.get("html_mechanic") == "deprecated"    # not recently committed => deprecates


def test_unused_experimental_not_promoted(isolated_state):
    rep = assess.assess(Canon.load())
    # With no telemetry, nothing earns promotion.
    assert all(t["to"] != "core" for t in rep["transitions"])


def test_structural_skill_spared_from_silence_deprecation(isolated_canon):
    # literature_scout is referenced by the GLOBAL workflow nodes (core_outward_trunk,
    # evidence_schema_flow) - structural. Offered-but-unused must NOT deprecate it:
    # it is phase-dormant (the literature rail is dormant in a figure phase), not unfit.
    for _ in range(9):
        telemetry.log("orient", domain="global", offered=["literature_scout"])
    rep = assess.assess(Canon.load())
    moves = {t["name"]: t["to"] for t in rep["transitions"]}
    assert moves.get("literature_scout") != "deprecated"
    # but a skill only name-dropped in a domain page (css_forge in the web axioms)
    # is NOT structural and still deprecates on the same silence.
    for _ in range(9):
        telemetry.log("orient", domain="web", offered=["css_forge"])
    rep2 = assess.assess(Canon.load())
    moves2 = {t["name"]: t["to"] for t in rep2["transitions"]}
    assert moves2.get("css_forge") == "deprecated"


def test_orient_dashboard_when_goal_unclear():
    c = Canon.load()
    # a blank goal surfaces the dashboard: active projects + the three activities,
    # so a session with no clear task picks a thread rather than guessing.
    bundle = c.orient("")
    assert "dashboard" in bundle
    dash = bundle["dashboard"]
    assert dash["active_projects"], "dashboard should list active projects"
    assert any(p["domain"] == "synthetics" for p in dash["active_projects"])
    assert set(dash["activities"]) == {"generating", "staging", "evaluating"}
    # a goal with a clear domain does NOT get the dashboard
    clear = c.orient("draft the adaptive normal form manuscript", domain="synthetics")
    assert "dashboard" not in clear


# ---------------------------------------------------------------------------
# dreamer (isolated canon copy outside the repo => independent git repo)
# ---------------------------------------------------------------------------


@pytest.fixture
def isolated_canon(tmp_path, monkeypatch):
    dst = tmp_path / "canon"
    shutil.copytree(canonmod._PACKAGED_CANON, dst)
    monkeypatch.setenv("LAPLACE_CANON", str(dst))
    monkeypatch.setenv("LAPLACE_STATE", str(tmp_path / "state"))
    telemetry.clear()
    # Deterministic lifecycle baseline, independent of whatever statuses the
    # shipped manifest happens to carry: the skills these tests exercise start
    # experimental. Committed as the repo's base so promotion/deprecation are
    # genuine transitions and a dreamer rollback has a target to revert to.
    seed = {"skills": {
        "latex_forge": {"status": "experimental"},
        "css_forge": {"status": "experimental"},
        "html_mechanic": {"status": "experimental"},  # non-structural (domain-only mention)
    }}
    (dst / "lifecycle.yaml").write_text(yaml.safe_dump(seed), encoding="utf-8")
    gitio.commit_all(dst, "seed: experimental baseline")
    canonmod.get_canon(fresh=True)
    yield dst
    telemetry.clear()


def test_dream_applies_transitions_commits_and_rolls_back(isolated_canon):
    for _ in range(6):
        telemetry.log("verify_check", skill="latex_forge", check="voice", passed=True)
    out = dreamer.dream(apply=True, forge=False)
    assert any(t["name"] == "latex_forge" and t["to"] == "core" for t in out["assessment"]["transitions"])
    assert out.get("commit")
    # status persisted through the lifecycle overlay
    assert canonmod.get_canon(fresh=True).skills["latex_forge"].status == "core"
    # rollback reverts
    rb = dreamer.rollback(out["commit"])
    assert rb["ok"]
    assert canonmod.get_canon(fresh=True).skills["latex_forge"].status == "experimental"


def test_dream_noop_without_evidence(isolated_canon):
    out = dreamer.dream(apply=True, forge=False)
    assert out["assessment"]["transitions"] == []
    assert out.get("commit") is None


def test_noop_dream_does_not_advance_boundary(isolated_canon):
    # An idle dream - empty window, nothing applied or forged - must NOT stamp the
    # boundary, or a later dream's window would be fragmented and friction orphaned.
    out0 = dreamer.dream(apply=True, forge=False)
    assert out0["assessment"]["events_seen"] == 0
    assert out0["boundary_advanced"] is False
    assert telemetry.last_dream_ts() is None
    # A productive dream (earns a transition) does advance the boundary.
    for _ in range(6):
        telemetry.log("verify_check", skill="latex_forge", check="voice", passed=True)
    out1 = dreamer.dream(apply=True, forge=False)
    assert out1["boundary_advanced"] is True
    assert telemetry.last_dream_ts() is not None


def test_preview_dream_does_not_consume_window(isolated_canon):
    # The leak that orphaned a real session: an apply=False preview over a FULL
    # window stamped the boundary, so the next apply=True dream saw zero events and
    # applied nothing. A preview (or any dream that applies no transition) must not
    # consume the window - the evidence has to survive for the next real dream.
    for _ in range(6):
        telemetry.log("verify_check", skill="latex_forge", check="voice", passed=True)
    preview = dreamer.dream(apply=False, forge=False)
    assert preview["assessment"]["events_seen"] == 6
    assert any(t["name"] == "latex_forge" for t in preview["assessment"]["transitions"])
    assert preview["boundary_advanced"] is False  # preview did NOT consume the window
    assert telemetry.last_dream_ts() is None
    # the real dream still sees the same evidence and applies it
    real = dreamer.dream(apply=True, forge=False)
    assert real["assessment"]["events_seen"] == 6
    assert real["boundary_advanced"] is True
    assert canonmod.get_canon(fresh=True).skills["latex_forge"].status == "core"


def test_events_since_filters_by_timestamp(isolated_state):
    telemetry.log("feedback", skill="latex_forge", signal=1)
    cutoff = telemetry.events()[-1]["ts"]  # timestamp of that event
    telemetry.log("feedback", skill="latex_forge", signal=1)  # after cutoff
    after = telemetry.events(since=cutoff)
    # only the event AT or AFTER cutoff survives; the one strictly before does not
    assert all(ev["ts"] >= cutoff for ev in after)
    assert len(after) >= 1


def test_dream_stamps_and_scopes_next_dream(isolated_canon):
    # Pre-dream events: earn a transition.
    for _ in range(6):
        telemetry.log("verify_check", skill="latex_forge", check="voice", passed=True)
    out1 = dreamer.dream(apply=True, forge=False)
    # dream_complete was logged — last_dream_ts is non-None.
    ts = telemetry.last_dream_ts()
    assert ts is not None
    # Second dream sees only events AFTER the first dream's stamp.
    # No new events => assess sees nothing => no transitions.
    out2 = dreamer.dream(apply=False, forge=False)
    assert out2["assessment"]["events_seen"] == 0


# ---------------------------------------------------------------------------
# render / client adapters
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("client", ["claude", "gemini", "codex", "antigravity"])
def test_sync_renders_each_client(client):
    res = render.sync(client, write=False)
    assert "laplace_orient" in res["content"]
    assert res["mcp_registration"]["mcpServers"]["laplace"]["command"] == "mcp-laplace"
    assert res["written"] is False


def test_sync_claude_has_skill_frontmatter():
    res = render.sync("claude", write=False)
    assert res["content"].startswith("---\nname: laplace\n")


# ---------------------------------------------------------------------------
# evidence skills (ledger completeness + alignment)
# ---------------------------------------------------------------------------

LEDGER = """# Evidence Ledger

## EPT-CLM-001: Strong claim about gain
**Claim:** The gain scales as sqrt(N).
**Derivation:** Appendix A.
**Literature:** Condorcet (1785).
**Numerical:** validate_gain().
**Status:** Proved.

## EPT-CLM-002: Weak claim about depth
**Claim:** Hierarchy depth must exceed three.
**Derivation:** Sketch only.
**Literature:** none
**Numerical:** TODO
**Status:** conjecture, gap remains.
"""


def test_evidence_ledger_flags_incomplete_and_weak(tmp_path):
    led = tmp_path / "evidence_ledger.md"
    led.write_text(LEDGER, encoding="utf-8")
    r = verify.run_backing("evidence_ledger", target=str(led))
    out = r["stdout"]
    assert "Records: 2" in out
    assert "EPT-CLM-002" in out  # weak/incomplete flagged
    assert "EPT-CLM-001" not in out.split("Incomplete")[-1].split("Non-affirmative")[0] \
        if "Incomplete" in out else True  # strong claim not in incomplete section
    assert r["returncode"] == 1  # incomplete records => exit 1


IDENTITY = """# Identity Ledger

## ID-001: Widget Theorem
**Kind:** coined-term
**Forms:** widget theorem, WT
**Role:** the core named result; must appear.
**Status:** load-bearing

## ID-002: Gadget framing
**Kind:** framing
**Forms:** gadget framing
**Role:** preferred cross-domain bridge.
**Status:** optional
"""


def test_identity_ledger_detects_drift(tmp_path):
    man = tmp_path / "identity_ledger.md"
    man.write_text(IDENTITY, encoding="utf-8")
    draft = tmp_path / "draft.tex"
    draft.write_text("This draft uses the gadget framing but never the core result.\n", encoding="utf-8")
    r = verify.run_backing("identity_ledger", target=str(draft), args=["--manifest", str(man)])
    assert "DRIFT" in r["stdout"]
    assert "ID-001" in r["stdout"]  # load-bearing Widget Theorem dropped
    assert r["returncode"] == 1


def test_identity_ledger_passes_when_present(tmp_path):
    man = tmp_path / "identity_ledger.md"
    man.write_text(IDENTITY, encoding="utf-8")
    draft = tmp_path / "draft.tex"
    draft.write_text("We prove the Widget Theorem via the gadget framing.\n", encoding="utf-8")
    r = verify.run_backing("identity_ledger", target=str(draft), args=["--manifest", str(man)])
    assert "No identity drift" in r["stdout"]
    assert r["returncode"] == 0


def test_evidence_alignment_tiers_and_finds_uncovered_goal(tmp_path):
    led = tmp_path / "evidence_ledger.md"
    led.write_text(LEDGER, encoding="utf-8")
    r = verify.run_backing(
        "evidence_alignment", target=str(led),
        args=["--goals", "gain, teleportation"],
    )
    out = r["stdout"]
    assert "1 strong, 0 partial, 1 weak" in out
    assert "gain" in out
    assert "teleportation" in out and "uncovered" in out.lower()
    assert (tmp_path / "support_map.md").exists()


def test_voice_corpus_reader_chunks_without_leaking_source(tmp_path):
    source = tmp_path / "source.txt"
    source.write_text(
        "\n\n".join(
            [
                "alpha private sentence " * 8,
                "beta private sentence " * 8,
                "gamma private sentence " * 8,
            ]
        ),
        encoding="utf-8",
    )
    cache = tmp_path / "cache"
    args = [
        "--source-handle", "TEST-SOURCE",
        "--register", "personal",
        "--source-kind", "unit",
        "--cache-root", str(cache),
        "--max-chars", "120",
        "--salt", "test-salt",
    ]

    first = verify.run_backing("voice_corpus_reader", target=str(source), args=args)
    assert first["returncode"] == 0
    assert "alpha private sentence" not in first["stdout"]
    assert "beta private sentence" not in first["stdout"]

    manifest_path = cache / "manifests" / "TEST-SOURCE.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["schema_version"] == "voice-corpus-chunks/0.1"
    assert manifest["source"]["source_handle"] == "TEST-SOURCE"
    assert manifest["source"]["register"] == "personal"
    assert len(manifest["chunks"]) == 3
    assert all((cache / "chunks" / "TEST-SOURCE" / f"{c['chunk_id']}.txt").exists() for c in manifest["chunks"])
    assert all(c["byte_end"] > c["byte_start"] for c in manifest["chunks"])

    first_ids = [c["chunk_id"] for c in manifest["chunks"]]
    first_source_hash = manifest["source"]["source_id_hash"]
    second = verify.run_backing("voice_corpus_reader", target=str(source), args=args)
    assert second["returncode"] == 0
    manifest2 = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert [c["chunk_id"] for c in manifest2["chunks"]] == first_ids
    assert manifest2["source"]["source_id_hash"] == first_source_hash


def test_voice_corpus_reader_can_prepare_restartable_queue(tmp_path):
    source = tmp_path / "source.txt"
    source.write_text(
        "\n\n".join(
            [
                "first private paragraph " * 8,
                "second private paragraph " * 8,
            ]
        ),
        encoding="utf-8",
    )
    cache = tmp_path / "cache"
    args = [
        "--source-handle", "QUEUE-SOURCE",
        "--register", "academic",
        "--source-kind", "unit",
        "--cache-root", str(cache),
        "--max-chars", "120",
        "--salt", "test-salt",
        "--reader-queue",
        "--reader-tasks", "voice,facts",
    ]

    result = verify.run_backing("voice_corpus_reader", target=str(source), args=args)
    assert result["returncode"] == 0
    assert "first private paragraph" not in result["stdout"]
    assert "Reader queue:" in result["stdout"]
    assert "Reader jobs: 4" in result["stdout"]

    queue_path = cache / "readers" / "QUEUE-SOURCE" / "queue.jsonl"
    queue_text = queue_path.read_text(encoding="utf-8")
    assert "first private paragraph" not in queue_text
    assert "second private paragraph" not in queue_text

    jobs = [json.loads(line) for line in queue_text.splitlines()]
    assert len(jobs) == 4
    assert {job["reader_task"] for job in jobs} == {"voice", "facts"}
    assert all(job["status"] == "pending" for job in jobs)
    assert all(job["attempts"] == 0 for job in jobs)
    assert all(job["output_path"].endswith(f".{job['reader_task']}.json") for job in jobs)


# ---------------------------------------------------------------------------
# watchdog: no handler may hang the caller (the assess-hang structural fix)
# ---------------------------------------------------------------------------


def test_run_backing_failure_records_returncode_and_stderr_tail(tmp_path, monkeypatch):
    """A non-zero backing-script exit must record returncode and stderr_tail in
    the emitted execute telemetry event, so the dreamer can self-diagnose
    without re-running the script.
    """
    import shutil

    # Build a minimal isolated canon with a skill whose backing script fails.
    dst = tmp_path / "canon"
    shutil.copytree(canonmod._PACKAGED_CANON, dst)
    monkeypatch.setenv("LAPLACE_CANON", str(dst))
    monkeypatch.setenv("LAPLACE_STATE", str(tmp_path / "state"))

    skill_dir = dst / "skills" / "failing_skill"
    skill_dir.mkdir(parents=True)
    (skill_dir / "SKILL.md").write_text(
        "---\ndescription: always fails\n---\n# Failing Skill\n",
        encoding="utf-8",
    )
    # A Python backing script that writes a known message to stderr and exits 2.
    backing = skill_dir / "run.py"
    backing.write_text(
        "import sys\n"
        "sys.stderr.write('something went wrong: detail here\\n')\n"
        "sys.exit(2)\n",
        encoding="utf-8",
    )

    # Register it in the manifest overlay.
    import yaml
    idx = yaml.safe_load((dst / "index.yaml").read_text(encoding="utf-8"))
    idx.setdefault("skills", {})["failing_skill"] = {
        "description": "always fails",
        "activity": "staging",
        "status": "experimental",
        "backing": "run.py",
        "tags": [],
    }
    (dst / "index.yaml").write_text(yaml.safe_dump(idx), encoding="utf-8")
    canonmod.get_canon(fresh=True)

    telemetry.clear()
    result = verify.run_backing("failing_skill")
    assert result["returncode"] == 2
    assert result["ok"] is False

    events = telemetry.events()
    exec_events = [e for e in events if e["phase"] == "execute" and e.get("skill") == "failing_skill"]
    assert exec_events, "no execute event recorded for failing skill"
    ev = exec_events[-1]
    assert ev["ok"] is False
    assert ev["returncode"] == 2, f"expected returncode=2 in event, got {ev}"
    assert "stderr_tail" in ev, "stderr_tail must be present in failure event"
    assert "something went wrong" in ev["stderr_tail"], (
        f"expected stderr content in stderr_tail, got {ev['stderr_tail']!r}"
    )
    # A successful run must NOT include returncode/stderr_tail.
    telemetry.clear()


def test_run_backing_success_does_not_record_failure_fields(tmp_path, monkeypatch):
    """A successful backing-script run must not include returncode/stderr_tail
    in the telemetry event - those fields are failure diagnostics only.
    """
    import shutil

    dst = tmp_path / "canon"
    shutil.copytree(canonmod._PACKAGED_CANON, dst)
    monkeypatch.setenv("LAPLACE_CANON", str(dst))
    monkeypatch.setenv("LAPLACE_STATE", str(tmp_path / "state"))

    skill_dir = dst / "skills" / "passing_skill"
    skill_dir.mkdir(parents=True)
    (skill_dir / "SKILL.md").write_text(
        "---\ndescription: always passes\n---\n# Passing Skill\n",
        encoding="utf-8",
    )
    backing = skill_dir / "run.py"
    backing.write_text("import sys\nprint('all good')\nsys.exit(0)\n", encoding="utf-8")

    import yaml
    idx = yaml.safe_load((dst / "index.yaml").read_text(encoding="utf-8"))
    idx.setdefault("skills", {})["passing_skill"] = {
        "description": "always passes",
        "activity": "staging",
        "status": "experimental",
        "backing": "run.py",
        "tags": [],
    }
    (dst / "index.yaml").write_text(yaml.safe_dump(idx), encoding="utf-8")
    canonmod.get_canon(fresh=True)

    telemetry.clear()
    result = verify.run_backing("passing_skill")
    assert result["returncode"] == 0
    assert result["ok"] is True

    events = telemetry.events()
    exec_events = [e for e in events if e["phase"] == "execute" and e.get("skill") == "passing_skill"]
    assert exec_events
    ev = exec_events[-1]
    assert ev["ok"] is True
    assert "returncode" not in ev, "success event must not carry returncode"
    assert "stderr_tail" not in ev, "success event must not carry stderr_tail"


def test_watchdog_passes_value_through():
    from mcp_gerard.laplace.watchdog import guard

    assert guard("ok", 5, lambda: {"v": 42}) == {"v": 42}


def test_watchdog_returns_timeout_dict_on_overrun():
    import time

    from mcp_gerard.laplace.watchdog import guard

    t0 = time.time()
    out = guard("slow", 0.3, lambda: time.sleep(10))
    elapsed = time.time() - t0
    assert out["timed_out"] is True
    assert "watchdog" in out["error"]
    assert elapsed < 2  # returned at the cap, did not wait out the 10s work


def test_watchdog_propagates_exceptions():
    from mcp_gerard.laplace.watchdog import guard

    def boom():
        raise ValueError("kaboom")

    with pytest.raises(ValueError, match="kaboom"):
        guard("boom", 5, boom)


def test_engine_git_is_pure_python_no_subprocess(tmp_path, monkeypatch):
    """Engine git is pure-Python (Dulwich): no subprocess, no git.exe, so no
    fsmonitor-daemon can inherit a pipe and deadlock it past its timeout.

    Proven two ways: gitio's source spawns no subprocess, and a commit still
    lands with subprocess.run forced to raise.
    """
    import inspect
    import subprocess as _sp

    # gitio shells nothing out - it is Dulwich end to end. (The word "subprocess"
    # in the module docstring is fine; what must be absent is any actual use.)
    src = inspect.getsource(gitio)
    assert "import subprocess" not in src, "gitio must not import subprocess"
    assert "subprocess.run" not in src and "subprocess.Popen" not in src
    assert "from dulwich" in src or "import dulwich" in src

    # A commit lands even with subprocess poisoned - there is no git.exe path.
    def _boom(*a, **k):
        raise AssertionError("engine git must not call subprocess.run")

    monkeypatch.setattr(_sp, "run", _boom)
    repo = tmp_path / "r"
    repo.mkdir()
    f = repo / "f.txt"
    f.write_text("hi", encoding="utf-8")
    sha = gitio.commit(repo, [f], "init")
    assert sha and gitio.head_sha(repo) == sha


def test_gitio_commit_log_revert_roundtrip(tmp_path):
    """The whole engine git surface: commit_all, commit, log_paths_since, revert."""
    repo = tmp_path / "r"
    repo.mkdir()
    f = repo / "lifecycle.yaml"
    f.write_text("v: 1\n", encoding="utf-8")
    sha1 = gitio.commit_all(repo, "seed")
    assert sha1

    f.write_text("v: 2\n", encoding="utf-8")
    sha2 = gitio.commit(repo, [f], "bump")
    assert sha2 and sha2 != sha1

    # the non-root commit's changed path shows up in the window
    paths = gitio.log_paths_since(repo, None)
    assert paths is not None and any(p.endswith("lifecycle.yaml") for p in paths)

    # reverting the bump restores the seed content
    ok, _detail = gitio.revert(repo, sha2)
    assert ok
    assert f.read_text(encoding="utf-8") == "v: 1\n"


# ---------------------------------------------------------------------------
# Item 1 - dreamer reads the backlog
# ---------------------------------------------------------------------------


def test_dream_exposes_backlog_headers(isolated_canon):
    """dream() result must carry a 'backlog' key with the section headings from
    AUTONOMY_BACKLOG.md, so deferred items resurface each cycle.
    """
    out = dreamer.dream(apply=False, forge=False)
    assert "backlog" in out, "dream() must include a 'backlog' key"
    backlog = out["backlog"]
    assert backlog.get("available") is True, f"backlog should be readable: {backlog}"
    sections = backlog.get("sections", [])
    assert sections, "backlog sections list must not be empty"
    # The backlog has at least an 'Engine behaviour' section.
    assert any("Engine behaviour" in s for s in sections), (
        f"Expected 'Engine behaviour' in backlog sections, got: {sections}"
    )
    # Open items should surface too.
    assert isinstance(backlog.get("open_items"), list)


# ---------------------------------------------------------------------------
# Item 2 - orient-guard for domain=null
# ---------------------------------------------------------------------------


def test_orient_null_domain_carries_candidate_domains():
    """When a goal is too ambiguous to infer a domain, the orient bundle must
    include candidate domains inline under 'domain_recovery', so the model has
    everything it needs to re-orient without a second round trip.
    """
    c = Canon.load()
    # A goal with no domain-specific language should yield domain=None.
    bundle = c.orient("do something")
    # If domain inference happened to find a match, skip this check.
    if bundle["domain"] is not None:
        pytest.skip("goal unexpectedly matched a domain; need a more ambiguous goal")
    assert "domain_recovery" in bundle, (
        "orient bundle must include 'domain_recovery' when domain=None"
    )
    recovery = bundle["domain_recovery"]
    assert "hint" in recovery, "domain_recovery must include a hint"
    candidates = recovery.get("candidate_domains", [])
    assert candidates, "domain_recovery must list candidate domains"
    # Each candidate must have a name.
    assert all("name" in cd for cd in candidates), (
        f"all candidate_domains entries must have a 'name' key: {candidates}"
    )


def test_orient_with_explicit_domain_has_no_recovery():
    """When a domain is supplied or cleanly inferred, no domain_recovery key
    should appear - it is only an aid for the underdetermined case.
    """
    c = Canon.load()
    # 'synthetics' is a known domain that should be inferred.
    bundle = c.orient(
        "lint the latex voice and check derivations for the phases of hierarchy"
    )
    assert bundle["domain"] == "synthetics"
    assert "domain_recovery" not in bundle, (
        "domain_recovery must not appear when domain is determined"
    )


# ---------------------------------------------------------------------------
# Item 3 - fitness must not read silence as dissatisfaction
# ---------------------------------------------------------------------------


def test_fitness_silence_does_not_deprecate_healthy_skill(isolated_state):
    """A skill with healthy usage and verify pass-rate but zero explicit feedback
    must not be down-ranked relative to its usage.  Absence of praise alone must
    never trigger a deprecate recommendation.
    """
    # Drive a healthy usage pattern: 6 verify passes, zero explicit feedback.
    for _ in range(6):
        telemetry.log("verify_check", skill="latex_forge", check="voice", passed=True)
    # No laplace_log calls - silence, not dissatisfaction.

    rep = assess.assess(Canon.load())
    skill_row = next((r for r in rep["skills"] if r["name"] == "latex_forge"), None)
    assert skill_row is not None

    # The skill must not be recommended for deprecation.
    assert skill_row["recommended_status"] != "deprecated", (
        f"latex_forge was recommended for deprecation despite healthy usage and "
        f"zero (not negative) feedback: {skill_row}"
    )
    # Fitness must be respectable - usage and quality should dominate.
    assert skill_row["fitness"] >= 0.5, (
        f"fitness {skill_row['fitness']} too low for a skill with healthy usage "
        f"and zero negative feedback"
    )


def test_absence_of_praise_never_triggers_deprecate_recommendation(isolated_state):
    """Even when a skill has usage but zero laplace_log calls, the assess report
    must not recommend deprecation - silence is not dissatisfaction.
    """
    # Enough usage to pass the PROMOTE_MIN_USES threshold.
    for _ in range(6):
        telemetry.log("execute", skill="result_foundry", kind="protocol")

    rep = assess.assess(Canon.load())
    skill_row = next((r for r in rep["skills"] if r["name"] == "result_foundry"), None)
    if skill_row is None:
        pytest.skip("result_foundry not in canon; pick any skill with zero feedback")

    transitions = {t["name"]: t["to"] for t in rep["transitions"]}
    assert transitions.get("result_foundry") != "deprecated", (
        "absence of positive feedback alone must not trigger a deprecate recommendation"
    )
