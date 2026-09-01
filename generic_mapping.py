import hashlib
import io
import json
import re
from datetime import datetime
from pathlib import Path


def clean_cell(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def split_word_line(words, gap=18):
    segments, current, previous_x1 = [], [], None
    for word in sorted(words, key=lambda item: item["x0"]):
        if previous_x1 is not None and word["x0"] - previous_x1 > gap and current:
            segments.append(current)
            current = []
        current.append(word)
        previous_x1 = word["x1"]
    if current:
        segments.append(current)
    return segments


PDF_HEADER_KEYWORDS = (
    "date", "particular", "description", "narration", "details", "memo",
    "cheque", "chq", "withdraw", "debit", "deposit", "credit", "balance",
    "amount", "reference", "utr",
)


def pdf_header_score(values):
    labels = [clean_cell(value).lower() for value in values]
    return sum(any(keyword in label for keyword in PDF_HEADER_KEYWORDS) for label in labels)


def header_token_kind(value):
    token = re.sub(r"[^a-z]", "", clean_cell(value).lower())
    if not token:
        return ""
    if token == "date" or token.endswith("date"):
        return "date"
    if any(part in token for part in ("particular", "description", "narration", "remark", "details", "memo")):
        return "particulars"
    if token.startswith(("chq", "cheque", "instrument")):
        return "instrument"
    if token.startswith(("withdraw", "debit")) or token == "dr":
        return "debit"
    if token.startswith(("deposit", "credit")) or token == "cr":
        return "credit"
    if token.endswith("balance"):
        return "balance"
    if token.startswith(("reference", "refno", "utr")):
        return "reference"
    if token in {"amount", "amt"}:
        return "amount"
    if token in {"drcr", "crdr", "direction", "type"}:
        return "direction"
    return ""


def word_header_anchors(words):
    anchors, previous_kind = [], ""
    for word in sorted(words, key=lambda item: item["x0"]):
        kind = header_token_kind(word["text"])
        if not kind:
            continue
        # "Debit Amount" and "Credit Amount" are one column heading.
        if kind == "amount" and previous_kind in {"debit", "credit"}:
            continue
        anchors.append(word["x0"])
        previous_kind = kind
    compact = []
    for anchor in anchors:
        if not compact or anchor - compact[-1] > 6:
            compact.append(anchor)
    return compact


def positioned_word_rows(page, table_anchors=None):
    words = page.extract_words(x_tolerance=2, y_tolerance=3) or []
    lines = []
    for word in sorted(words, key=lambda item: (round(item["top"] / 3), item["x0"])):
        line_key = round(word["top"] / 3)
        if not lines or lines[-1][0] != line_key:
            lines.append((line_key, []))
        lines[-1][1].append(word)

    anchors = list(table_anchors or [])
    if len(anchors) < 3:
        best_anchors = []
        for _, line_words in lines:
            candidate = word_header_anchors(line_words)
            if 3 <= len(candidate) <= 15 and len(candidate) > len(best_anchors):
                best_anchors = candidate
        if len(best_anchors) >= 3:
            anchors = best_anchors

    rows = []
    if len(anchors) >= 3:
        for _, line_words in lines:
            cells = [""] * len(anchors)
            for word in sorted(line_words, key=lambda item: item["x0"]):
                column = 0
                for index, anchor in enumerate(anchors):
                    if word["x0"] >= anchor - 4:
                        column = index
                    else:
                        break
                cells[column] = (cells[column] + " " + word["text"]).strip()
            if sum(bool(clean_cell(cell)) for cell in cells) > 1:
                rows.append(cells)
        return rows

    for _, line_words in lines:
        cells = [" ".join(word["text"] for word in segment) for segment in split_word_line(line_words)]
        if len(cells) > 1:
            rows.append(cells)
    return rows


def extract_pdf_grid(raw, password=""):
    try:
        import pdfplumber
    except ImportError as exc:
        raise ValueError("PDF table support is not installed. Run Install_Requirements.bat.") from exc
    try:
        pdf = pdfplumber.open(io.BytesIO(raw), password=password or None)
    except Exception as exc:
        message = str(exc).lower()
        if "password" in message or "decrypt" in message:
            raise PermissionError("PDF_PASSWORD_REQUIRED") from exc
        raise
    grid = []
    with pdf:
        last_anchors = None
        for page in pdf.pages:
            table_signatures = set()
            found_tables = page.find_tables() or []
            table_anchors, anchor_score = None, 0
            # A bank PDF often exposes its page heading and transaction body as
            # separate tables. Keeping only the largest table can therefore keep
            # the repeated heading while silently dropping every transaction.
            for found_table in found_tables:
                table = found_table.extract() or []
                for row in table:
                    cleaned = [clean_cell(cell) for cell in row]
                    if any(cleaned):
                        grid.append(cleaned)
                        table_signatures.add(tuple(cleaned))
                for row_index, row in enumerate(table):
                    score = pdf_header_score(row)
                    if score <= anchor_score or row_index >= len(found_table.rows):
                        continue
                    cells = found_table.rows[row_index].cells
                    if cells and all(cell is not None for cell in cells):
                        table_anchors = [cell[0] for cell in cells]
                        anchor_score = score
            if table_anchors:
                last_anchors = table_anchors
            # Always inspect positioned page text as well. Some banks draw only
            # the headings as a table and place transactions as loose PDF words.
            # Multi-page statements may omit the heading on later pages, so reuse
            # the most recent column positions instead of collapsing Date and
            # Narration into one cell.
            for row in positioned_word_rows(page, table_anchors or last_anchors):
                cleaned = [clean_cell(cell) for cell in row]
                if tuple(cleaned) not in table_signatures:
                    grid.append(cleaned)
    if not grid:
        raise ValueError("No text table was detected. This scanned PDF requires OCR.")
    width = max(len(row) for row in grid)
    return [row + [""] * (width - len(row)) for row in grid]


def format_fingerprint(filename, grid, header_row):
    suffix = Path(filename).suffix.lower()
    headers = [clean_cell(value).lower() for value in grid[int(header_row)]]
    return hashlib.sha256((suffix + "|" + "|".join(headers)).encode("utf-8")).hexdigest()[:24]


def parse_date(value):
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    text = clean_cell(value)

    for fmt in (
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d-%m-%y",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%d.%m.%Y",
        "%d.%m.%y",
        "%d %b %Y",
        "%d %B %Y",
    ):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    return text


def number(value):
    cleaned = re.sub(r"[^\d.\-]", "", clean_cell(value).replace(",", ""))
    try:
        return round(float(cleaned), 2)
    except ValueError:
        return 0.0


def apply_mapping(grid, mapping, header_row, bank_ledger, classify):
    def column(row, key, default=""):
        index = mapping.get(key)
        if index in (None, "", -1):
            return default
        index = int(index)
        return row[index] if index < len(row) else default
    output = []
    running_balance = None
    debit_column = mapping.get("debit")
    credit_column = mapping.get("credit")
    shared_amount_column = (
        debit_column
        if debit_column not in (None, "", -1) and str(debit_column) == str(credit_column)
        else None
    )
    for source in grid[int(header_row) + 1:]:
        if not any(clean_cell(value) for value in source):
            continue
        raw_date = column(source, "date")
        parsed_date = parse_date(raw_date)
        particulars = clean_cell(column(source, "particulars"))
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", parsed_date) or not particulars:
            continue
        debit = 0 if shared_amount_column is not None else number(column(source, "debit"))
        credit = 0 if shared_amount_column is not None else number(column(source, "credit"))
        negative_debit = abs(debit) if debit < 0 else 0
        negative_credit = abs(credit) if credit < 0 else 0
        debit = round(max(debit, 0) + negative_credit, 2)
        credit = round(max(credit, 0) + negative_debit, 2)
        if not debit and not credit and (
            mapping.get("amount") not in (None, "", -1) or shared_amount_column is not None
        ):
            value = number(source[int(shared_amount_column)]) if shared_amount_column is not None else number(column(source, "amount"))
            direction = clean_cell(column(source, "direction")).upper()
            if direction.startswith("C") or direction in {"CR", "CREDIT", "DEPOSIT"}:
                credit = value
            else:
                debit = value
        balance = number(column(source, "balance"))
        if not debit and not credit and balance and running_balance is not None:
            change = round(balance - running_balance, 2)
            debit = abs(change) if change < 0 else 0
            credit = change if change > 0 else 0
        running_balance = balance or running_balance
        base = {
            "date": parsed_date,
            "value_date": parse_date(column(source, "value_date", raw_date)),
            "instrument": clean_cell(column(source, "instrument")),
            "particulars": particulars,
            "reference": clean_cell(column(source, "reference")),
            "debit": debit,
            "credit": credit,
            "balance": balance,
            "balance_available": mapping.get("balance") not in (None, "", -1),
        }
        output.append(base)
    if not output:
        raise ValueError(
            "No transaction rows matched this mapping. Select the header row above "
            "the actual transactions and confirm the Date and Particulars columns."
        )
    if mapping.get("balance") not in (None, "", -1) and len(output) > 1:
        newest_first = output[0]["date"] > output[-1]["date"]
        chronological = sorted(
            enumerate(output),
            key=lambda item: (
                item[1]["date"],
                -item[0] if newest_first else item[0],
            ),
        )
        previous_balance = chronological[0][1]["balance"]
        previous_row = chronological[0][1]
        duplicate_ids = set()
        for _, row in chronological[1:]:
            change = round(row["balance"] - previous_balance, 2)
            if not change:
                # Some PDF tables repeat the same transaction: first with a
                # shortened narration and again with the full narration, while
                # the running balance stays unchanged. Keep one transaction and
                # preserve the more complete narration.
                if len(row.get("particulars", "")) > len(previous_row.get("particulars", "")):
                    previous_row["particulars"] = row["particulars"]
                    previous_row["reference"] = row.get("reference") or previous_row.get("reference", "")
                    previous_row["instrument"] = row.get("instrument") or previous_row.get("instrument", "")
                duplicate_ids.add(id(row))
                continue
            row["debit"] = abs(change) if change < 0 else 0
            row["credit"] = change if change > 0 else 0
            previous_balance = row["balance"]
            previous_row = row
        output = [row for row in output if id(row) not in duplicate_ids]
    return [classify(row, bank_ledger) for row in output]


def load_profiles(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_profiles(path, profiles):
    Path(path).write_text(json.dumps(profiles, ensure_ascii=False, indent=2), encoding="utf-8")
