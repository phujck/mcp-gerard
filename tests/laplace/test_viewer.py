"""The live viewer server: serves the typed graph + persists node comments.

No browser needed - the server is exercised in-process over loopback. The Canvas
app it serves is verified separately with the preview tools.
"""

from __future__ import annotations

import json
import threading
import urllib.parse
import urllib.request
from http.server import ThreadingHTTPServer

from mcp_gerard.laplace import viewer


def _start(source, path):
    httpd = ThreadingHTTPServer(("127.0.0.1", 0), viewer.make_handler(source, path))
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd, httpd.server_address[1]


def _get(base, path):
    with urllib.request.urlopen(base + path, timeout=10) as r:
        return r.read()


def _post(base, path, payload):
    req = urllib.request.Request(
        base + path, data=json.dumps(payload).encode(), headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req, timeout=10) as r:
        return json.loads(r.read())


def test_viewer_serves_canon_graph_and_persists_comments(tmp_path, monkeypatch):
    monkeypatch.setenv("LAPLACE_STATE", str(tmp_path / "state"))
    httpd, port = _start("canon", None)
    base = f"http://127.0.0.1:{port}"
    try:
        html = _get(base, "/").decode()
        assert "Laplace" in html and "<canvas" in html

        g = json.loads(_get(base, "/graph?source=canon"))
        art = g.get("artifact", g)
        assert art["nodes"], "canon graph should have nodes"

        r = _post(base, "/comment", {"node": "skill:manuscript_spine", "text": "looks good", "source": "canon"})
        assert r["ok"]
        cs = json.loads(_get(base, "/comments?source=canon"))
        assert cs["skill:manuscript_spine"][0]["text"] == "looks good"
        assert cs["skill:manuscript_spine"][0]["ts"]  # timestamped
    finally:
        httpd.shutdown()


def test_viewer_serves_blueprint_and_comments_beside_it(tmp_path, monkeypatch):
    monkeypatch.setenv("LAPLACE_STATE", str(tmp_path / "state"))
    bp = tmp_path / "blueprint.md"
    bp.write_text(
        "# Blueprint: t\n\n## sec:a | A | ring:core | status:stub\n"
        "- claim: x | result:R1 | status:established\n",
        encoding="utf-8",
    )
    httpd, port = _start("blueprint", str(bp))
    base = f"http://127.0.0.1:{port}"
    try:
        q = urllib.parse.quote(str(bp))
        g = json.loads(_get(base, f"/graph?source=blueprint&path={q}"))
        art = g.get("artifact", g)
        kinds = {n["kind"] for n in art["nodes"]}
        assert "section" in kinds and "claim" in kinds

        _post(base, "/comment", {"node": "sec:a", "text": "tighten the frame", "source": "blueprint", "path": str(bp)})
        # comment lands in a project-side sidecar, never in the canon
        assert (tmp_path / "blueprint_comments.json").exists()
    finally:
        httpd.shutdown()
