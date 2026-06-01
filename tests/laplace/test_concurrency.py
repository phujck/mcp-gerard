"""Concurrency safety for canon mutation under multiple agents.

Three agents (Claude, Codex, Gemini-on-Antigravity) may drive the engine at
once. These tests pin the two guarantees that makes safe:

* the canon write lock serialises a critical section *across processes* (not
  just threads), so no mutation clobbers another;
* lifecycle writes are atomic, so a lock-free reader never sees torn YAML;
* a contended lock returns a clean busy result fast, never hangs.
"""

from __future__ import annotations

import subprocess
import sys
import threading
import time
from pathlib import Path
from types import SimpleNamespace

import yaml


# ---------------------------------------------------------------------------
# Cross-process serialisation. Real subprocesses (sys.executable -c ...) drive
# the same lock - this sidesteps the multiprocessing/pickling/import-name
# fragility of pytest workers on Windows `spawn`, while genuinely testing the
# cross-process guarantee against the installed package.
# ---------------------------------------------------------------------------
_WORKER = """
import os, sys
state, counter, iters = sys.argv[1], sys.argv[2], int(sys.argv[3])
os.environ["LAPLACE_STATE"] = state
from pathlib import Path
from mcp_gerard.laplace.locking import canon_lock
p = Path(counter)
for _ in range(iters):
    with canon_lock(timeout=120):
        n = int(p.read_text(encoding="utf-8") or "0")
        p.write_text(str(n + 1), encoding="utf-8")
"""


def test_canon_lock_serialises_across_processes(tmp_path):
    """N processes each do a non-atomic read-modify-write under the lock.

    Without mutual exclusion the interleaved increments lose updates and the
    total is < N*M. With the cross-process lock it is exactly N*M.
    """
    state_dir = tmp_path / "state"
    state_dir.mkdir()
    counter = tmp_path / "counter.txt"
    counter.write_text("0", encoding="utf-8")

    n_procs, iters = 4, 25
    procs = [
        subprocess.Popen([sys.executable, "-c", _WORKER, str(state_dir), str(counter), str(iters)])
        for _ in range(n_procs)
    ]
    for p in procs:
        assert p.wait(timeout=180) == 0, "worker process failed"

    assert int(counter.read_text(encoding="utf-8")) == n_procs * iters


# ---------------------------------------------------------------------------
# Atomic lifecycle writes: a reader parsing the file in a tight loop while
# writers rewrite it must never see invalid YAML.
# ---------------------------------------------------------------------------
def test_save_lifecycle_atomic_never_torn(tmp_path, monkeypatch):
    # Production faithful: writers serialise on the canon lock, readers stay
    # lock-free. The atomic os.replace is what guarantees a lock-free reader
    # never sees a torn, mid-write overlay.
    monkeypatch.setenv("LAPLACE_STATE", str(tmp_path / "state"))
    from mcp_gerard.laplace import dreamer
    from mcp_gerard.laplace.locking import canon_lock

    canon = SimpleNamespace(root=tmp_path)
    path = tmp_path / "lifecycle.yaml"
    dreamer._save_lifecycle(canon, {"skills": {}})

    stop = threading.Event()
    errors: list[str] = []

    def writer(i: int) -> None:
        # A chunky payload so a torn read would be obvious as a YAML error.
        data = {"skills": {f"s{i}": {"status": "experimental", "history": [{"at": str(k)} for k in range(60)]}}}
        while not stop.is_set():
            with canon_lock(timeout=30):
                dreamer._save_lifecycle(canon, data)

    def reader() -> None:
        while not stop.is_set():
            try:
                yaml.safe_load(path.read_text(encoding="utf-8"))
            except FileNotFoundError:
                pass  # mid-swap absence is tolerable; torn content is not
            except Exception as exc:  # noqa: BLE001
                errors.append(repr(exc))
                return
            # Realistic cadence: the canon loader reads on a fingerprint change,
            # not in a zero-gap loop. A tiny yield keeps heavy overlap without
            # pinning the file open 100% of the time (which would be unphysical).
            time.sleep(0.001)

    threads = [threading.Thread(target=writer, args=(i,)) for i in range(3)]
    threads += [threading.Thread(target=reader) for _ in range(2)]
    for t in threads:
        t.start()
    time.sleep(1.0)
    stop.set()
    for t in threads:
        t.join(timeout=10)

    assert not errors, f"reader saw torn/invalid YAML: {errors}"
    assert yaml.safe_load(path.read_text(encoding="utf-8")) is not None


# ---------------------------------------------------------------------------
# A contended lock returns a structured busy result fast, never hangs.
# ---------------------------------------------------------------------------
def test_lock_busy_returns_fast(tmp_path, monkeypatch):
    monkeypatch.setenv("LAPLACE_STATE", str(tmp_path))
    from mcp_gerard.laplace.locking import CanonBusy, canon_lock

    with canon_lock(timeout=5):
        start = time.monotonic()
        raised = False
        try:
            with canon_lock(timeout=0.5):
                pass
        except CanonBusy:
            raised = True
        elapsed = time.monotonic() - start

    assert raised, "second acquisition should have raised CanonBusy"
    assert elapsed < 3.0, f"busy-return took {elapsed:.1f}s, should be ~0.5s"
