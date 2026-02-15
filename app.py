from __future__ import annotations

import json
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from xlsx_store import CharacterAppStore

ROOT = Path(__file__).parent


class AppHandler(SimpleHTTPRequestHandler):
    store = CharacterAppStore(ROOT)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/state":
            self._send_json(self.store.build_state())
            return
        if parsed.path == "/":
            self.path = "/templates/index.html"
        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/action":
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length) if length else b"{}")
            result = self.store.apply_action(payload)
            self._send_json(result)
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def _send_json(self, payload: dict):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


if __name__ == "__main__":
    server = ThreadingHTTPServer(("0.0.0.0", 8000), AppHandler)
    print("Serveur démarré sur http://localhost:8000")
    server.serve_forever()
