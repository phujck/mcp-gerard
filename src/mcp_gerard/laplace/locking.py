"""Cross-process write lock for canon mutations.

Three agents - Claude, Codex, Gemini-on-Antigravity - may drive the engine at
once. A canon mutation (a lifecycle transition, a forged skill, a rollback)
rewrites shared files (``lifecycle.yaml``) and commits the canon repo. Without
serialisation, two concurrent dreams read the same lifecycle, mutate
independently, and the second write clobbers the first - an observed lost
update. This lock makes the read-assess-mutate-commit critical section
single-writer *across processes*.

Two invariants:

* **Reads never take it.** ``orient``/``assess``/``skill``/``graph``/``verify``
  stay lock-free against the mtime-fingerprinted canon cache. Only mutation
  serialises, so the constant-traffic read path can never be blocked by a write.
* **The lock file lives under the state dir** (``LAPLACE_STATE``), never inside
  the canon repo, so it is never itself a tracked, committed, or synced artifact.

Acquisition is bounded: on contention it raises :class:`CanonBusy` rather than
blocking forever, and the caller turns that into a structured "busy, retry"
result. The tool watchdog is the outer backstop; this timeout sits well inside
it.
"""

from __future__ import annotations

import contextlib
from typing import Iterator

from filelock import FileLock, Timeout

from mcp_gerard.laplace import telemetry as _telemetry


class CanonBusy(RuntimeError):
    """The canon write lock could not be acquired within the timeout."""


def lock_path() -> str:
    """The single canon-write lock, under the state dir (never in the canon)."""
    return str(_telemetry.state_dir() / "canon.lock")


@contextlib.contextmanager
def canon_lock(timeout: float = 20.0) -> Iterator[None]:
    """Hold the one canon-write lock for a mutate+commit critical section.

    Raises :class:`CanonBusy` if the lock cannot be acquired within ``timeout``
    seconds, so a contended caller returns a clean busy result instead of
    hanging. A single lock with no nesting means no lock-ordering and so no
    deadlock. Exceptions from the wrapped body propagate unchanged (the lock is
    released on every path).
    """
    lock = FileLock(lock_path(), timeout=timeout)
    try:
        lock.acquire()
    except Timeout as exc:
        raise CanonBusy(
            f"canon write lock busy after {timeout:.0f}s - another agent is "
            "mutating the canon, retry shortly"
        ) from exc
    try:
        yield
    finally:
        lock.release()
