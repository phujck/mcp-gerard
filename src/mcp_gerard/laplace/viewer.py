"""Local web viewer for the live preview network (canon | blueprint | manuscript).

Serves a single-file Canvas viewer plus a small JSON API over ``graph.render``,
so the author can drill the same typed graph from macro (sections) to atom
(claim <-> evidence), and comment on any node. Three verbs in one pane:
*understand* (the picture, at any scale), *insert* (a comment, or the next
elicited answer), and the engine *reconciles*.

Pure stdlib ``http.server``. It NEVER spawns git - versioning is Dulwich; this
is a read-and-comment surface - so it cannot reintroduce the fsmonitor deadlock
class. Node comments are persisted to a project-side sidecar (next to the
blueprint), never into the canon, so the viewer is not a fourth concurrent canon
writer (see the canon write lock).
"""

from __future__ import annotations

import argparse
import json
import os
import tempfile
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from mcp_gerard.laplace import graph as _graph

_APP = Path(__file__).resolve().parent / "viewer_app.html"


def graph_json(source: str, path: str | None, focus=None, depth=2, kinds=None) -> dict:
    """Render the typed graph for the requested source. No git, ever."""
    if source == "blueprint" and path:
        return _graph.render("json", focus=focus, depth=depth, kinds=kinds, blueprint=path)
    if source == "manuscript" and path:
        return _graph.render("json", focus=focus, depth=depth, kinds=kinds, manuscript=path)
    return _graph.render("json", focus=focus, depth=depth, kinds=kinds)


def _comments_path(source: str, path: str | None) -> Path:
    if path:
        p = Path(path)
        return p.with_name(p.stem + "_comments.json")  # project-side, beside the blueprint
    from mcp_gerard.laplace import telemetry

    return telemetry.state_dir() / "canon_comments.json"


def _load_comments(cp: Path) -> dict:
    try:
        return json.loads(cp.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_comments(cp: Path, data: dict) -> None:
    cp.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(cp.parent), prefix=cp.name + ".", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(json.dumps(data, indent=2, ensure_ascii=False))
        os.replace(tmp, cp)
    finally:
        if Path(tmp).exists():
            Path(tmp).unlink()


def make_handler(source: str, path: str | None):
    class Handler(BaseHTTPRequestHandler):
        def _send(self, code: int, body, ctype: str = "application/json") -> None:
            data = body if isinstance(body, bytes) else str(body).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self) -> None:  # noqa: N802
            u = urlparse(self.path)
            q = parse_qs(u.query)
            if u.path in ("/", "/index.html"):
                self._send(200, _APP.read_bytes(), "text/html; charset=utf-8")
            elif u.path == "/meta":
                self._send(200, json.dumps({"source": source, "path": path}))
            elif u.path == "/graph":
                src = q.get("source", [source])[0]
                pth = q.get("path", [path])[0]
                focus = q.get("focus", [None])[0]
                depth = int(q.get("depth", ["2"])[0])
                kinds = q.get("kinds", [None])[0]
                res = graph_json(src, pth, focus, depth, kinds.split(",") if kinds else None)
                self._send(200, json.dumps(res))
            elif u.path == "/comments":
                src = q.get("source", [source])[0]
                pth = q.get("path", [path])[0]
                self._send(200, json.dumps(_load_comments(_comments_path(src, pth))))
            else:
                self._send(404, json.dumps({"error": "not found"}))

        def do_POST(self) -> None:  # noqa: N802
            u = urlparse(self.path)
            if u.path != "/comment":
                self._send(404, json.dumps({"error": "not found"}))
                return
            length = int(self.headers.get("Content-Length", 0))
            try:
                payload = json.loads(self.rfile.read(length) or b"{}")
            except ValueError:
                self._send(400, json.dumps({"error": "bad json"}))
                return
            src = payload.get("source", source)
            pth = payload.get("path", path)
            node = payload.get("node", "")
            cp = _comments_path(src, pth)
            data = _load_comments(cp)
            data.setdefault(node, []).append(
                {
                    "text": payload.get("text", ""),
                    "label": payload.get("label", ""),
                    "ts": datetime.now(timezone.utc).isoformat(),
                }
            )
            _save_comments(cp, data)
            self._send(200, json.dumps({"ok": True, "node": node, "count": len(data[node])}))

        def log_message(self, *a) -> None:  # keep the console quiet
            pass

    return Handler


def serve(source: str = "canon", path: str | None = None, port: int = 8765, host: str = "127.0.0.1") -> None:
    httpd = ThreadingHTTPServer((host, port), make_handler(source, path))
    print(f"laplace viewer: http://{host}:{port}  (source={source} path={path})", flush=True)
    httpd.serve_forever()


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Laplace live graph viewer.")
    p.add_argument("--source", default="canon", choices=["canon", "blueprint", "manuscript"])
    p.add_argument("--path", default=None, help="blueprint.md or manuscript .tex (for those sources)")
    p.add_argument("--port", type=int, default=8765)
    p.add_argument("--host", default="127.0.0.1")
    args = p.parse_args(argv)
    serve(args.source, args.path, args.port, args.host)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
