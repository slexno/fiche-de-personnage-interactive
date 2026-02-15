from __future__ import annotations

import argparse
import json
import os
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from xlsx_store import CharacterAppStore

ROOT = Path(__file__).parent


class AppHandler(SimpleHTTPRequestHandler):
    store = CharacterAppStore(ROOT)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

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


def run_server(preferred_port: int = 80):
    tried = []
    for port in [preferred_port, 80, 8000, 8080, 5000, 8001, 8888]:
        if port in tried:
            continue
        tried.append(port)
        try:
            server = ThreadingHTTPServer(("0.0.0.0", port), AppHandler)
            if port == 80:
                print("Serveur démarré sur http://localhost")
            else:
                print(f"Serveur démarré sur http://localhost:{port}")
            server.serve_forever()
            return
        except OSError:
            continue
    raise OSError("Impossible de démarrer le serveur: ports 80/8000/8080/5000/8001/8888 indisponibles")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fiche de personnage interactive")
    parser.add_argument("--port", type=int, default=int(os.getenv("PORT", "80")), help="Port HTTP (80 par défaut)")
    args = parser.parse_args()
    run_server(args.port)
