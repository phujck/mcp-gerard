"""Pure-Python git for the engine - no subprocess, no daemon, no fsmonitor.

The engine used to shell ``git.exe`` and capture its output through an OS pipe.
On Windows that deadlocked: Claude Code's background ``git-fsmonitor--daemon``
inherited the pipe's write handle, so the read never saw EOF and the call hung
past its own timeout - ``laplace_assess`` froze for ~30 minutes at a time.

Dulwich removes the entire class of bug by construction. There is no child
process, no pipe, and no daemon, so nothing can be inherited and nothing can
hang. Every canon mutation - a lifecycle write, a forged skill, a rollback -
versions through here.

The engine only ever commits a small set of canon files and never merges, so
the surface is deliberately tiny: ``ensure_repo``, ``commit``, ``commit_all``,
``head_sha``, ``log_paths_since``, and a path-scoped ``revert``. Pushing to a
remote stays out of this hot path (an explicit, occasional action), so Dulwich's
network/auth surface is never touched here.
"""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from dulwich import porcelain
from dulwich.repo import Repo

# A stable author so the canon history is legibly the engine's own.
_AUTHOR = b"Laplace Engine <laplace@localhost>"


def _b(s: str | bytes) -> bytes:
    return s.encode("utf-8") if isinstance(s, str) else s


def _abspaths(root: Path, paths) -> list[str]:
    out = []
    for p in paths:
        pp = Path(p)
        out.append(str(pp if pp.is_absolute() else (root / pp)))
    return out


def ensure_repo(root) -> Repo:
    """Open the git repo at ``root``, initialising one if absent."""
    root = Path(root)
    if (root / ".git").exists():
        return Repo(str(root))
    root.mkdir(parents=True, exist_ok=True)
    return porcelain.init(str(root))


def commit(root, paths, message: str) -> str | None:
    """Stage ``paths`` and commit. Returns the new sha, or None if nothing staged.

    Mirrors the old ``git add <paths> && git diff --cached --quiet || git commit``:
    if staging produced no change against HEAD, there is nothing to commit.
    """
    root = Path(root)
    repo = ensure_repo(root)
    porcelain.add(repo, paths=_abspaths(root, paths))
    st = porcelain.status(repo)
    if not any(st.staged.values()):
        return None
    sha = porcelain.commit(repo, message=_b(message), author=_AUTHOR, committer=_AUTHOR)
    return sha.decode() if isinstance(sha, (bytes, bytearray)) else str(sha)


def commit_all(root, message: str) -> str | None:
    """Stage every tracked-or-new file under ``root`` (excluding .git) and commit."""
    root = Path(root)
    repo = ensure_repo(root)
    files = [str(p) for p in root.rglob("*") if p.is_file() and ".git" not in p.parts]
    if files:
        porcelain.add(repo, paths=files)
    sha = porcelain.commit(repo, message=_b(message), author=_AUTHOR, committer=_AUTHOR)
    return sha.decode() if isinstance(sha, (bytes, bytearray)) else str(sha)


def head_sha(root) -> str | None:
    """The current HEAD sha, or None if the repo has no commits / does not exist."""
    try:
        return Repo(str(Path(root))).head().decode()
    except Exception:  # noqa: BLE001 - any repo/HEAD problem is "unknown"
        return None


def _parse_since(since: str | None) -> int:
    """ISO timestamp -> epoch seconds. None / unparseable -> 24h ago (a safe default)."""
    if since:
        try:
            dt = datetime.fromisoformat(since)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return int(dt.timestamp())
        except ValueError:
            pass
    return int(datetime.now(timezone.utc).timestamp()) - 86400


def log_paths_since(root, since: str | None) -> set[str] | None:
    """Repo-relative paths changed by non-root commits since ``since``.

    Replaces ``git log --since --min-parents=1 --name-only --format=``. Returns
    None when git is unavailable (caller falls back to "no grace"), an empty set
    when the repo has no qualifying commits.
    """
    root = Path(root)
    if not (root / ".git").exists():
        return None
    try:
        repo = Repo(str(root))
    except Exception:  # noqa: BLE001
        return None
    try:
        head = repo.head()
    except KeyError:
        return set()  # no commits yet
    cutoff = _parse_since(since)
    paths: set[str] = set()
    try:
        for entry in repo.get_walker(include=[head], since=cutoff):
            commit_obj = entry.commit
            if len(commit_obj.parents) < 1:
                continue  # --min-parents=1: skip the root commit
            changes = entry.changes()
            for ch in changes:
                for c in (ch if isinstance(ch, list) else [ch]):  # merges -> list of lists
                    for te in (c.new, c.old):
                        if te is not None and te.path:
                            paths.add(te.path.decode("utf-8", "replace").replace("\\", "/"))
    except Exception:  # noqa: BLE001 - a history read problem is "unknown"
        return None
    return paths


def _resolve_ref(repo: Repo, ref: str):
    if ref in ("HEAD", b"HEAD"):
        return repo.head()
    rb = ref.encode() if isinstance(ref, str) else ref
    if rb in repo.object_store:
        return rb
    try:
        return repo.refs[rb]
    except KeyError:
        for sha in repo.object_store:  # short-sha fallback (canon repos are small)
            if sha.startswith(rb):
                return sha
        raise KeyError(ref)


def revert(root, ref: str) -> tuple[bool, str]:
    """Path-scoped inverse-restore of one commit, then commit "Revert <ref>".

    Dulwich has no porcelain revert. The engine never merges and only ever
    commits a small file set, so reverting is exact: for every path the target
    commit changed, restore the parent's version (or delete a file the commit
    added), then record the inverse as a new commit. Returns (ok, sha-or-error).
    """
    from dulwich.diff_tree import tree_changes
    from dulwich.object_store import tree_lookup_path

    root = Path(root)
    try:
        repo = Repo(str(root))
        cid = _resolve_ref(repo, ref)
        commit_obj = repo[cid]
        if not commit_obj.parents:
            return False, "cannot revert the root commit"
        parent = repo[commit_obj.parents[0]]
        touched: list[str] = []
        for ch in tree_changes(repo.object_store, parent.tree, commit_obj.tree):
            pe = ch.new if (ch.new and ch.new.path) else ch.old
            if pe is None or not pe.path:
                continue
            fp = root / pe.path.decode("utf-8", "replace")
            try:
                _mode, blob_sha = tree_lookup_path(repo.get_object, parent.tree, pe.path)
                fp.parent.mkdir(parents=True, exist_ok=True)
                fp.write_bytes(repo[blob_sha].data)
            except KeyError:
                if fp.exists():
                    fp.unlink()  # the commit added it; the inverse removes it
            touched.append(str(fp))
        if touched:
            porcelain.add(repo, paths=touched)
        sha = porcelain.commit(repo, message=_b(f"Revert {ref}"), author=_AUTHOR, committer=_AUTHOR)
        return True, (sha.decode() if isinstance(sha, (bytes, bytearray)) else str(sha))
    except Exception as exc:  # noqa: BLE001
        return False, str(exc)
