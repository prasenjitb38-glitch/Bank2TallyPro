from __future__ import annotations

import base64
import hashlib
import hmac
import io
import json
import os
import platform
import re
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    import winreg
except ImportError:
    winreg = None


TRIAL_CREDITS = 50
SHORT_CODE_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
SHORT_CODE_SECRET = base64.b64decode("m2fVx6U0r8bL4QnZp7eS3jK9cW1aY5dH+oTgIuNqP0E=")
SELLER_MOBILE = "9508773595"
SHORT_PLANS = {
    1: ("pack", 200),
    2: ("pack", 300),
    3: ("pack", 500),
    4: ("pack", 1000),
    5: ("yearly", 2500),
}
PUBLIC_N = "wDAplu5U6iIUL2/mUa6EHMjjZDCU5bK7vEu3ognwsxNtGiatk3oeRUHoYv7FMpum3QdtaaiFBUvDNppVyKOhuH8fHXalftwjOZzDq1ObECHe1bT5U01tsAKLkblJp/FL2+XsNS61N1wh/YUxZ8urDzDzFxTXGBE48ilN2kBGW/0jxpg4qIjsosjcCXWw/gOwU0ej94ZgfLLhfsnwQEj2/HcIgmMDcIVFkQs2wsiI0NfmfWid2Vm/LNsQSIcu5FRrwztbgvOcJCBa0YRYuORPUmVlQRo2FltbKUCt40/JyqxqyK3eCRgn4Wmpgr28p2QZ6QjWKRsywGK06nmFrKIxQQ=="
PUBLIC_E = "AQAB"
SHA256_DER_PREFIX = bytes.fromhex("3031300d060960864801650304020105000420")


def _b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode().rstrip("=")


def _unb64url(text: str) -> bytes:
    return base64.urlsafe_b64decode(text + "=" * (-len(text) % 4))


def _machine_source() -> str:
    if winreg:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
                return str(winreg.QueryValueEx(key, "MachineGuid")[0])
        except OSError:
            pass
    return f"{platform.node()}|{uuid.getnode()}|{platform.machine()}"


def device_id() -> str:
    digest = hashlib.sha256(("Bank2Tally|" + _machine_source()).encode()).hexdigest().upper()
    return "-".join(digest[i:i + 4] for i in range(0, 20, 4))


class LicenseStore:
    def __init__(self, data_dir: Path):
        self.path = data_dir / "license_state.json"
        self.device = device_id()
        self.state = self._load()

    def _default(self):
        return {
            "device_id": self.device,
            "trial_total": TRIAL_CREDITS,
            "trial_used": 0,
            "paid_credits": 0,
            "yearly_credits": 0,
            "yearly_expiry": "",
            "activated_keys": [],
            "processed_files": {},
        }

    def _load(self):
        try:
            state = json.loads(self.path.read_text(encoding="utf-8"))
            if state.get("device_id") == self.device:
                return {**self._default(), **state}
        except (OSError, ValueError):
            pass
        state = self._default()
        self._save(state)
        return state

    def _save(self, state=None):
        if state is not None:
            self.state = state
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temp = self.path.with_suffix(".tmp")
        temp.write_text(json.dumps(self.state, indent=2), encoding="utf-8")
        os.replace(temp, self.path)

    def _yearly_available(self):
        expiry = self.state.get("yearly_expiry", "")
        if not expiry:
            return 0
        try:
            if date.today() > date.fromisoformat(expiry):
                return 0
        except ValueError:
            return 0
        return max(0, int(self.state.get("yearly_credits", 0)))

    def status(self):
        trial = max(0, int(self.state["trial_total"]) - int(self.state["trial_used"]))
        paid = max(0, int(self.state.get("paid_credits", 0)))
        yearly = self._yearly_available()
        return {
            "device_id": self.device,
            "trial_remaining": trial,
            "paid_remaining": paid,
            "yearly_remaining": yearly,
            "remaining": trial + paid + yearly,
            "yearly_expiry": self.state.get("yearly_expiry", ""),
            "plan": "Yearly" if yearly else ("Credit Pack" if paid else "Trial"),
        }

    def already_processed(self, digest):
        return digest in self.state.get("processed_files", {})

    def can_charge(self, credits):
        return self.status()["remaining"] >= credits

    def charge(self, credits, digest, filename):
        credits = max(1, int(credits))
        if self.already_processed(digest):
            return {"charged": 0, "duplicate": True, **self.status()}
        if not self.can_charge(credits):
            raise ValueError(f"Not enough credits. This file needs {credits} credit(s).")
        remaining = credits
        trial = max(0, int(self.state["trial_total"]) - int(self.state["trial_used"]))
        take = min(trial, remaining)
        self.state["trial_used"] = int(self.state["trial_used"]) + take
        remaining -= take
        yearly = self._yearly_available()
        take = min(yearly, remaining)
        self.state["yearly_credits"] = yearly - take
        remaining -= take
        paid = max(0, int(self.state.get("paid_credits", 0)))
        take = min(paid, remaining)
        self.state["paid_credits"] = paid - take
        remaining -= take
        self.state.setdefault("processed_files", {})[digest] = {
            "name": filename, "credits": credits, "processed_at": datetime.now().isoformat(timespec="seconds")
        }
        self._save()
        return {"charged": credits, "duplicate": False, **self.status()}

    def activate(self, key, customer_mobile=""):
        compact = "".join(ch for ch in key.upper() if ch.isalnum())
        if len(compact) == 20 and all(ch in SHORT_CODE_ALPHABET for ch in compact):
            return self._activate_short(compact, customer_mobile)
        try:
            payload_text, signature_text = key.strip().split(".", 1)
            payload_bytes = _unb64url(payload_text)
            signature = _unb64url(signature_text)
            self._verify(payload_bytes, signature)
            payload = json.loads(payload_bytes)
        except Exception as exc:
            raise ValueError("Invalid license key.") from exc
        if payload.get("device_id", "").upper() != self.device:
            raise ValueError("This license key belongs to a different computer.")
        license_id = payload.get("license_id", "")
        if not license_id or license_id in self.state["activated_keys"]:
            raise ValueError("This license key has already been used.")
        credits = int(payload.get("credits", 0))
        if credits <= 0:
            raise ValueError("License has no credits.")
        if payload.get("type") == "yearly":
            expiry = date.fromisoformat(payload["expiry"])
            if expiry < date.today():
                raise ValueError("This yearly license has expired.")
            self.state["yearly_credits"] = int(self.state.get("yearly_credits", 0)) + credits
            current = self.state.get("yearly_expiry", "")
            self.state["yearly_expiry"] = max(current, expiry.isoformat())
        else:
            self.state["paid_credits"] = int(self.state.get("paid_credits", 0)) + credits
        self.state["activated_keys"].append(license_id)
        self._save()
        return self.status()

    def _activate_short(self, compact, customer_mobile):
        number = 0
        for char in compact:
            number = number * 32 + SHORT_CODE_ALPHABET.index(char)
        payload_number = number >> 68
        received_mac = number & ((1 << 68) - 1)
        version = payload_number >> 28
        plan_id = (payload_number >> 25) & 7
        if version != 1 or plan_id not in SHORT_PLANS:
            raise ValueError("Invalid activation code.")
        mobile = re.sub(r"\D", "", customer_mobile)
        message = payload_number.to_bytes(4, "big") + f"{self.device}|{mobile}|{SELLER_MOBILE}".encode()
        expected_mac = int.from_bytes(hmac.new(SHORT_CODE_SECRET, message, hashlib.sha256).digest()[:9], "big") >> 4
        if not hmac.compare_digest(received_mac.to_bytes(9, "big"), expected_mac.to_bytes(9, "big")):
            raise ValueError("Activation code does not match this Device ID or Mobile Number.")
        license_id = "SHORT-" + compact
        if license_id in self.state["activated_keys"]:
            raise ValueError("This activation code has already been used.")
        kind, credits = SHORT_PLANS[plan_id]
        if kind == "yearly":
            self.state["yearly_credits"] = int(self.state.get("yearly_credits", 0)) + credits
            expiry = date.today() + timedelta(days=365)
            self.state["yearly_expiry"] = max(self.state.get("yearly_expiry", ""), expiry.isoformat())
        else:
            self.state["paid_credits"] = int(self.state.get("paid_credits", 0)) + credits
        self.state["activated_keys"].append(license_id)
        self._save()
        return self.status()

    @staticmethod
    def _verify(payload, signature):
        n = int.from_bytes(base64.b64decode(PUBLIC_N), "big")
        e = int.from_bytes(base64.b64decode(PUBLIC_E), "big")
        size = (n.bit_length() + 7) // 8
        encoded = pow(int.from_bytes(signature, "big"), e, n).to_bytes(size, "big")
        expected = SHA256_DER_PREFIX + hashlib.sha256(payload).digest()
        padding = b"\x00\x01" + b"\xff" * (size - len(expected) - 3) + b"\x00" + expected
        if encoded != padding:
            raise ValueError("Bad signature")


def file_digest(raw):
    return hashlib.sha256(raw).hexdigest()


def credit_cost(filename, raw, transaction_count=0, password=""):
    if Path(filename).suffix.lower() == ".pdf":
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(raw))
        if reader.is_encrypted:
            if not password or not reader.decrypt(password):
                raise ValueError("PDF password is required before credits can be calculated.")
        return max(1, len(reader.pages))
    return max(1, (int(transaction_count) + 24) // 25)
