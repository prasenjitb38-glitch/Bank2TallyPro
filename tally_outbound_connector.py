"""Outbound Bank2Tally Connector — no Cloudflare tunnel or inbound URL."""
from __future__ import annotations
import base64, json, os, time, urllib.parse, urllib.request
from pathlib import Path
from scanned_bank_ocr import parse_scanned_bank_pdf

SERVER_URL = os.environ.get("CONNECTOR_SERVER_URL", "https://bank2tally-suite.onrender.com").rstrip("/")
CONNECTOR_ID = os.environ.get("TALLY_CONNECTOR_ID", "primary")
TOKEN = os.environ.get("TALLY_CONNECTOR_TOKEN", "")
TALLY_URL = os.environ.get("TALLY_URL", "http://127.0.0.1:9000")
OCR_SCRIPT = Path(__file__).resolve().parent / "windows_ocr.ps1"

def call(path, method="GET", data=None, timeout=35):
    raw = None if data is None else json.dumps(data).encode("utf-8")
    request = urllib.request.Request(SERVER_URL + path, data=raw, method=method,
        headers={"X-Connector-Token": TOKEN, "Content-Type": "application/json"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))

def execute(command):
    if command["kind"] == "health":
        try:
            send_tally("<ENVELOPE><HEADER><TALLYREQUEST>Export</TALLYREQUEST></HEADER></ENVELOPE>", 10)
            return {"ok": True, "reachable": True, "detail": "TallyPrime is reachable through the outbound Connector."}
        except Exception as exc:
            return {"ok": True, "reachable": False, "detail": f"TallyPrime is not reachable: {exc}"}
    if command["kind"] == "tally":
        return {"ok": True, "body": send_tally(command["xml"], command.get("timeout", 30))}
    if command["kind"] == "ocr":
        rows, opening, closing = parse_scanned_bank_pdf(base64.b64decode(command["data"]), OCR_SCRIPT)
        return {"ok": True, "rows": rows, "opening": opening, "closing": closing}
    return {"ok": False, "error": "Unknown Connector command."}

def send_tally(xml, timeout):
    request = urllib.request.Request(TALLY_URL, data=xml.encode("utf-8"), method="POST",
        headers={"Content-Type": "text/xml; charset=utf-8"})
    with urllib.request.urlopen(request, timeout=min(max(int(timeout), 1), 120)) as response:
        return response.read().decode("utf-8", errors="replace")

if not TOKEN:
    raise SystemExit("Set TALLY_CONNECTOR_TOKEN before starting the Connector.")
print("Bank2Tally Connector started — connected securely without any tunnel.")
retry = 1
while True:
    try:
        command = call(f"/api/connector/poll?id={urllib.parse.quote(CONNECTOR_ID)}").get("command")
        retry = 1
        if not command:
            continue
        try:
            result = execute(command)
        except Exception as exc:
            result = {"ok": False, "error": str(exc)}
        call("/api/connector/result", "POST", {"connector_id": CONNECTOR_ID, "command_id": command["id"], "result": result})
    except Exception as exc:
        print(f"[Connector] {exc}; retrying in {retry}s")
        time.sleep(retry)
        retry = min(retry * 2, 30)
