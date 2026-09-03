"""Local Windows bridge for Bank2Tally cloud deployments.

Run this on the computer where TallyPrime is open and Windows OCR is available.
Expose it only through a private/secured tunnel and set BRIDGE_TOKEN.
"""
from __future__ import annotations

import base64
import json
import os
import subprocess
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from scanned_bank_ocr import parse_scanned_bank_pdf


HOST = os.environ.get("BRIDGE_HOST", "127.0.0.1")
PORT = int(os.environ.get("BRIDGE_PORT", "9010"))
TOKEN = os.environ.get("BRIDGE_TOKEN", "")
TALLY_URL = os.environ.get("TALLY_URL", "http://127.0.0.1:9000")
OCR_SCRIPT = Path(__file__).resolve().parent / "windows_ocr.ps1"


def json_bytes(value):
    return json.dumps(value, ensure_ascii=False).encode("utf-8")


class BridgeHandler(BaseHTTPRequestHandler):
    def _auth(self):
        return bool(TOKEN) and self.headers.get("X-Bridge-Token", "") == TOKEN

    def _send(self, status, payload):
        raw = json_bytes(payload)
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Bridge-Token")
        self.end_headers()
        self.wfile.write(raw)

    def do_OPTIONS(self):
        self._send(204, {})

    def do_GET(self):
        if self.path == "/health":
            self._send(200, {"ok": True, "service": "Bank2Tally Local Bridge"})
            return
        self._send(404, {"error": "Not found"})

    def do_POST(self):
        if not self._auth():
            self._send(401, {"error": "Bridge authentication failed."})
            return
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        try:
            if self.path == "/tally":
                request = urllib.request.Request(
                    TALLY_URL, data=body,
                    headers={"Content-Type": "text/xml; charset=utf-8"}, method="POST"
                )
                with urllib.request.urlopen(request, timeout=120) as response:
                    result = response.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/xml; charset=utf-8")
                self.send_header("Content-Length", str(len(result)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(result)
                return
            if self.path == "/ocr":
                payload = json.loads(body.decode("utf-8"))
                raw = base64.b64decode(payload.get("data", ""))
                rows, opening, closing = parse_scanned_bank_pdf(raw, OCR_SCRIPT)
                self._send(200, {"rows": rows, "opening": opening, "closing": closing})
                return
            self._send(404, {"error": "Not found"})
        except Exception as exc:
            self._send(400, {"error": str(exc)})

    def log_message(self, fmt, *args):
        print("[Bridge] " + (fmt % args))


if __name__ == "__main__":
    if not TOKEN:
        raise SystemExit("Set BRIDGE_TOKEN before starting the connector.")
    server = ThreadingHTTPServer((HOST, PORT), BridgeHandler)
    print(f"Bank2Tally Local Bridge listening on http://{HOST}:{PORT}")
    print(f"Forwarding TallyPrime requests to {TALLY_URL}")
    server.serve_forever()
