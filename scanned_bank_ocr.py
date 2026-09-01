from __future__ import annotations

import json
import re
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path


DATE_RE = re.compile(r"^(\d{2})[-/.](\d{2})[-/.](\d{4})$")


def _ocr_number(value):
    original = str(value or "").upper()
    # Do not turn ordinary words such as “Opening” into the number 0 merely
    # because OCR read the initial O as a digit.
    if not re.search(r"\d", original):
        return None
    text = original.replace("O", "0").replace(",", "")
    text = re.sub(r"[^0-9.]", "", text)
    if not text:
        return None
    try:
        if "." in text:
            return round(float(text), 2)
        # OCR commonly removes the decimal point from the Balance column.
        return round(int(text) / 100, 2) if len(text) >= 4 else float(text)
    except ValueError:
        return None


def _date(value):
    raw = str(value or "").strip().replace("—", "-")
    match = DATE_RE.match(raw)
    if not match:
        # Bandhan's scanned statement uses dates such as 30-Jun-2025.
        # Windows OCR can confuse a few month glyphs (Ju1, xug, Scp, 1k).
        textual = re.match(r"^(\d{1,2})[-/.]([A-Za-z0-9]{3})[-/.](\d{4})$", raw)
        if not textual:
            return ""
        day, month, year = textual.groups()
        month = {
            "JU1": "JUL", "XUG": "AUG", "SC P": "SEP", "SCP": "SEP",
            "1K": "DEC", "0EC": "DEC",
        }.get(month.upper().replace(" ", ""), month.upper())
        raw = f"{day}-{month}-{year}"
        try:
            return datetime.strptime(raw, "%d-%b-%Y").strftime("%Y-%m-%d")
        except ValueError:
            return ""
    try:
        return datetime.strptime("-".join(match.groups()), "%d-%m-%Y").strftime("%Y-%m-%d")
    except ValueError:
        return ""


def _is_balance_word(word, width):
    x = float(word.get("x", 0))
    # Bandhan places the running balance in the far-right column.  The
    # preceding credit column can also contain money values, so keep this
    # band narrow enough not to treat credits as balances.
    if not (width * 0.84 <= x <= width * 0.98):
        return False
    raw = str(word.get("text", "")).replace(",", "").strip()
    digits = re.sub(r"\D", "", raw)
    # Axis statements have a narrow Init. Br column immediately after Balance
    # containing values such as 271/400. It must never become a balance row.
    if x > width * 0.85 and "." not in raw and len(digits) <= 3:
        return False
    return _ocr_number(raw) is not None


def _run_windows_ocr(image_path, script_path):
    completed = subprocess.run(
        [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(script_path), "-Path", str(image_path),
        ],
        capture_output=True,
        timeout=90,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if completed.returncode:
        error_text = completed.stderr.decode("utf-8-sig", errors="replace")
        raise ValueError((error_text or "Windows OCR failed.").strip())
    output_text = completed.stdout.decode("utf-8-sig", errors="replace")
    output_text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", output_text)
    try:
        result = json.loads(output_text.strip() or "[]")
    except json.JSONDecodeError as exc:
        raise ValueError("Windows OCR returned unreadable data.") from exc
    return result if isinstance(result, list) else [result]


def _page_rows(words, width, page_index):
    date_words = []
    for word in words:
        parsed = _date(word.get("text"))
        if parsed and float(word.get("x", 0)) < width * 0.25:
            date_words.append((float(word.get("y", 0)), parsed))
    date_words.sort()
    if not date_words:
        return [], None

    opening = None
    opening_words = [
        word for word in words
        if "OPENING" in str(word.get("text", "")).upper()
    ]
    if opening_words:
        opening_y = min(float(word.get("y", 0)) for word in opening_words)
        candidates = []
        for word in words:
            x, y = float(word.get("x", 0)), float(word.get("y", 0))
            if abs(y - opening_y) <= 100:
                value = _ocr_number(word.get("text"))
                if value is not None:
                    # Opening balance is in the first summary column on
                    # Bandhan statements, unlike the running balance column.
                    if float(word.get("x", 0)) < width * 0.30 or _is_balance_word(word, width):
                        candidates.append((abs(y - opening_y), value))
        if candidates:
            opening = sorted(candidates)[0][1]

    # Running balance is present on every transaction row, while a faint or
    # skewed scan can occasionally hide a date. Use balance positions as the
    # row anchors, then attach the date found in that row (or carry it forward).
    first_date_y, last_date_y = date_words[0][0], date_words[-1][0]
    balance_anchors = []
    for word in words:
        x, y = float(word.get("x", 0)), float(word.get("y", 0))
        if not _is_balance_word(word, width):
            continue
        if not (first_date_y - 40 <= y <= last_date_y + 40):
            continue
        value = _ocr_number(word.get("text"))
        if value is not None:
            balance_anchors.append((y, value))
    balance_anchors.sort()

    output = []
    current_date = ""
    for index, (anchor_y, balance) in enumerate(balance_anchors):
        lower = (balance_anchors[index - 1][0] + anchor_y) / 2 if index else anchor_y - 35
        upper = (anchor_y + balance_anchors[index + 1][0]) / 2 if index + 1 < len(balance_anchors) else anchor_y + 40
        in_band = [
            word for word in words
            if lower <= float(word.get("y", 0)) < upper
        ]
        band_dates = [
            _date(word.get("text")) for word in in_band
            if float(word.get("x", 0)) < width * 0.25 and _date(word.get("text"))
        ]
        if band_dates:
            current_date = band_dates[0]
        particulars_words = [
            word for word in in_band
            if width * 0.22 <= float(word.get("x", 0)) < width * 0.59
        ]
        particulars_words.sort(key=lambda item: (round(float(item.get("y", 0)) / 6), float(item.get("x", 0))))
        particulars = " ".join(str(word.get("text", "")).strip() for word in particulars_words).strip()

        output.append({
            "date": current_date,
            "value_date": current_date,
            "instrument": "",
            "particulars": particulars or "Scanned bank transaction",
            "reference": "",
            "debit": 0.0,
            "credit": 0.0,
            "balance": balance,
            "balance_available": True,
            "_page": page_index,
            "_y": anchor_y,
        })
    return output, opening


def parse_scanned_bank_pdf(raw, script_path):
    try:
        import pypdfium2 as pdfium
    except ImportError as exc:
        raise ValueError(
            "Scanned PDF support needs pypdfium2. Run Install_Requirements.bat once."
        ) from exc

    document = pdfium.PdfDocument(raw)
    all_rows = []
    opening = None
    with tempfile.TemporaryDirectory(prefix="bank2tally_ocr_") as temp_dir:
        for page_index in range(len(document)):
            page = document[page_index]
            bitmap = page.render(scale=3.2)
            image = bitmap.to_pil()
            image_path = Path(temp_dir) / f"page-{page_index + 1}.png"
            image.save(image_path)
            words = _run_windows_ocr(image_path, script_path)
            page_rows, page_opening = _page_rows(words, image.width, page_index)
            if opening is None and page_opening is not None:
                opening = page_opening
            all_rows.extend(page_rows)

    if not all_rows or opening is None:
        raise ValueError(
            "OCR could not identify the statement rows or opening balance. "
            "Use a clearer scan (straight pages, at least 150 DPI)."
        )

    previous = opening
    previous_date = ""
    for row in all_rows:
        if row["date"]:
            previous_date = row["date"]
        else:
            row["date"] = previous_date
            row["value_date"] = previous_date
        if not row["date"]:
            raise ValueError("OCR could not identify the date of the first transaction row.")
        change = round(row["balance"] - previous, 2)
        row["debit"] = abs(change) if change < 0 else 0.0
        row["credit"] = change if change > 0 else 0.0
        previous = row["balance"]
        row.pop("_page", None)
        row.pop("_y", None)
    return all_rows, round(opening, 2), round(previous, 2)
