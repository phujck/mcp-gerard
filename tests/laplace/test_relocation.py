"""Canon relocation: seed bootstrap + isolation from the code repo.

The canon is relocatable to a user-local path via LAPLACE_CANON so the
self-refining canon versions itself in its own repo, invisible to the mcp-gerard
code repo and shared across the three agents. These tests pin:

* a fresh LAPLACE_CANON target is seeded from the packaged canon on first load;
* a real canon is never overwritten by the seed;
* the dreamer's commit lands in the external canon's own git repo, not the code
  repo.
"""

from __future__ import annotations

import subprocess

import pytest


@pytest.fixture
def fresh_canon_env(tmp_path, monkeypatch):
    """Point the canon + state dirs at tmp and clear the load cache."""
    canon_dir = tmp_path / "canon"
    state_dir = tmp_path / "state"
    monkeypatch.setenv("LAPLACE_CANON", str(canon_dir))
    monkeypatch.setenv("LAPLACE_STATE", str(state_dir))
    from mcp_gerard.laplace import canon as canon_mod

    canon_mod._CANON_CACHE.clear()
    yield canon_dir, state_dir
    canon_mod._CANON_CACHE.clear()


def test_seed_bootstrap_copies_packaged_canon(fresh_canon_env):
    canon_dir, _ = fresh_canon_env
    from mcp_gerard.laplace import canon as canon_mod

    assert not (canon_dir / "index.yaml").exists()  # empty before first load
    c = canon_mod.get_canon(fresh=True)

    assert (canon_dir / "index.yaml").exists()  # seeded from the package
    assert (canon_dir / "skills").is_dir()
    assert c.skills, "seeded canon should expose its skills"
    assert "manuscript_spine" in c.skills  # a skill we know is in the packaged canon


def test_seed_never_overwrites_a_real_canon(fresh_canon_env):
    canon_dir, _ = fresh_canon_env
    from mcp_gerard.laplace import canon as canon_mod

    # A pre-existing canon (its own index.yaml) must be left untouched.
    canon_dir.mkdir(parents=True)
    (canon_dir / "index.yaml").write_text("version: 1\nskills: {}\n", encoding="utf-8")
    sentinel = "# SENTINEL - do not overwrite\nversion: 1\nskills: {}\n"
    (canon_dir / "index.yaml").write_text(sentinel, encoding="utf-8")

    canon_mod.get_canon(fresh=True)

    assert (canon_dir / "index.yaml").read_text(encoding="utf-8") == sentinel
    assert not (canon_dir / "skills").exists()  # not seeded over


def test_dream_commit_lands_in_external_canon_repo(fresh_canon_env):
    canon_dir, _ = fresh_canon_env
    from mcp_gerard.laplace import canon as canon_mod
    from mcp_gerard.laplace import dreamer

    canon_mod.get_canon(fresh=True)  # seeds the external canon
    target = canon_dir / "lifecycle.yaml"
    target.write_text("# test\nskills: {}\n", encoding="utf-8")

    sha = dreamer._commit(canon_dir, [target], "test: external canon commit")

    assert sha, "commit should have landed and returned a sha"
    assert (canon_dir / ".git").exists(), "external canon should hold its own repo"
    log = subprocess.run(
        ["git", "-c", "core.fsmonitor=false", "-C", str(canon_dir), "log", "--oneline", "-1"],
        capture_output=True,
        text=True,
    )
    assert "external canon commit" in log.stdout
