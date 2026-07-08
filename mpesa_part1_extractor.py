# =============================================================================
# PART 1 — M-PESA PDF STATEMENT EXTRACTOR & CSV GENERATOR
# =============================================================================
# Requirements: pip install pdfplumber pandas
# =============================================================================

import re
import sys
import pdfplumber
import pandas as pd
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
PDF_FOLDER = Path(r"D:\Downloads\MPESA")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — RAW TEXT EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────
def extract_raw_text(pdf_path: str) -> list[str]:
    pages_text = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)
            print(f"[INFO] PDF opened — {total} page(s) detected.")
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text(x_tolerance=2, y_tolerance=3)
                if text:
                    pages_text.append(text)
                else:
                    pages_text.append(f"__BLANK_PAGE_{i}__")
                    print(f"[WARN] Page {i} returned no text — may be image-only.")
    except Exception as exc:
        print(f"[ERROR] pdfplumber failed: {exc}")
        print("[INFO] Attempting fallback with PyMuPDF …")
        try:
            import fitz
            doc = fitz.open(pdf_path)
            total = len(doc)
            print(f"[INFO] PyMuPDF opened — {total} page(s) detected.")
            for i, page in enumerate(doc, 1):
                text = page.get_text("text")
                pages_text.append(text if text.strip() else f"__BLANK_PAGE_{i}__")
            doc.close()
        except ImportError:
            sys.exit("[FATAL] Neither pdfplumber nor PyMuPDF could open the file.")
        except Exception as exc2:
            sys.exit(f"[FATAL] PyMuPDF fallback also failed: {exc2}")
    return pages_text


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — TEXT CLEANING
# ─────────────────────────────────────────────────────────────────────────────
NOISE_PATTERNS = [
    re.compile(r"Page\s+\d+\s+of\s+\d+", re.IGNORECASE),
    re.compile(r"M-PESA\s+Statement", re.IGNORECASE),
    re.compile(r"Safaricom\s+PLC", re.IGNORECASE),
    re.compile(r"Statement\s+Period.*", re.IGNORECASE),
    re.compile(r"Customer\s+Name.*", re.IGNORECASE),
    re.compile(r"Mobile\s+Number.*", re.IGNORECASE),
    re.compile(r"Account\s+Number.*", re.IGNORECASE),
    re.compile(r"Receipt\s+No\.?\s+Completion\s+Time", re.IGNORECASE),
    re.compile(r"Details\s+Status\s+Paid\s+In", re.IGNORECASE),
    re.compile(r"Withdrawn\s+Balance", re.IGNORECASE),
    re.compile(r"^-+$"),
    re.compile(r"^\s*$"),
    re.compile(r"^_{3,}$"),
    re.compile(r"^={3,}$"),
    re.compile(r"This\s+is\s+a\s+system\s+generated", re.IGNORECASE),
    re.compile(r"Confidential.*", re.IGNORECASE),
    re.compile(r"^\s*Continued\s+on\s+next\s+page\s*$", re.IGNORECASE),
]


def clean_page_lines(raw_text: str) -> list[str]:
    cleaned = []
    for line in raw_text.splitlines():
        line = line.strip()
        if not line or line.startswith("__BLANK"):
            continue
        if any(pat.search(line) for pat in NOISE_PATTERNS):
            continue
        cleaned.append(line)
    return cleaned


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — TRANSACTION PARSING
# ─────────────────────────────────────────────────────────────────────────────
RECEIPT_RE  = re.compile(r"^([A-Z0-9]{8,14})\s+")
DATETIME_RE = re.compile(r"(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})")
AMOUNT_RE   = re.compile(r"-?[\d,]+\.\d{2}")
STATUS_RE   = re.compile(r"\b(Completed|Failed|Reversed|Pending)\b", re.IGNORECASE)


def parse_amount_str(text: str) -> float:
    try:
        return float(text.replace(",", ""))
    except (ValueError, AttributeError):
        return 0.0


def parse_lines_into_transactions(all_lines: list[str]) -> list[dict]:
    raw_rows = []
    current = None
    for line in all_lines:
        if RECEIPT_RE.match(line):
            if current is not None:
                raw_rows.append(current)
            current = line
        else:
            if current is not None:
                current += " " + line
    if current:
        raw_rows.append(current)
    print(f"[INFO] Raw transaction rows assembled: {len(raw_rows)}")
    transactions = []
    for row in raw_rows:
        txn = parse_single_row(row)
        if txn:
            transactions.append(txn)
    return transactions


def parse_single_row(row: str) -> dict | None:
    m_receipt = RECEIPT_RE.match(row)
    if not m_receipt:
        return None
    receipt_no = m_receipt.group(1)

    m_dt = DATETIME_RE.search(row)
    if not m_dt:
        return None
    raw_date = m_dt.group(1)
    raw_time = m_dt.group(2)

    try:
        dt_obj = pd.to_datetime(f"{raw_date} {raw_time}", format="%Y-%m-%d %H:%M:%S")
    except Exception:
        dt_obj = pd.NaT

    m_status = STATUS_RE.search(row)
    status = m_status.group(1).capitalize() if m_status else "Unknown"

    amount_strings = AMOUNT_RE.findall(row)
    paid_in_val   = 0.0
    withdrawn_val = 0.0
    balance_val   = 0.0

    if amount_strings:
        balance_val = parse_amount_str(amount_strings[-1])
        prior = amount_strings[:-1]
        for amt_str in prior:
            val = parse_amount_str(amt_str)
            if val < 0:
                if withdrawn_val == 0.0:
                    withdrawn_val = val
            else:
                if paid_in_val == 0.0 and val > 0:
                    paid_in_val = val

    after_dt    = row[m_dt.end():]
    description = STATUS_RE.sub("", after_dt)
    description = AMOUNT_RE.sub("", description)
    description = re.sub(r"\s{2,}", " ", description).strip(" -,.")

    txn_type = infer_transaction_type(description, paid_in_val, withdrawn_val)

    return {
        "receipt_no":       receipt_no,
        "date":             raw_date,
        "time":             raw_time,
        "datetime":         dt_obj,
        "details":          description,
        "transaction_type": txn_type,
        "status":           status,
        "paid_in":          paid_in_val,
        "withdrawn":        withdrawn_val,
        "balance":          balance_val,
    }


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — TRANSACTION TYPE INFERENCE
# ─────────────────────────────────────────────────────────────────────────────
TYPE_KEYWORDS = [
    ("Fuliza / Overdraft",     ["fuliza", "overdraft", "overdraw", "od loan"]),
    ("Airtime Purchase",       ["airtime"]),
    ("Data Bundle",            ["bundle", "data bundle", "safaricom data"]),
    ("Paybill Payment",        ["pay bill", "paybill"]),
    ("Merchant Payment",       ["merchant payment", "merchant pay"]),
    ("Send Money",             ["send money", "customer transfer", "transfer to",
                                "payment to small business", "customer payment"]),
    ("Receive Money",          ["funds received", "receive", "business payment"]),
    ("International Transfer", ["international", "forex", "money transfer"]),
    ("Unit Trust / Savings",   ["unit trust", "ziidi", "mmf", "invest", "withdraw from"]),
    ("KCB Loan",               ["kcb m-pesa loan", "kcb loan"]),
    ("ATM / Agent Withdrawal", ["withdrawal at agent", "withdrawal charge",
                                "customer withdrawal"]),
    ("Bank Transfer",          ["equity paybill", "co-operative bank",
                                "kcb paybill", "absa bank", "lipa na kcb"]),
    ("Reversal",               ["reversal", "reversed"]),
    ("Charges / Fees",         ["charge", "fee"]),
    ("Savings Contribution",   ["savings contribution", "term loan"]),
    ("KPLC / Utility",         ["kplc", "kisumu water", "utility"]),
]


def infer_transaction_type(description: str, paid_in: float, withdrawn: float) -> str:
    desc_lower = description.lower()
    for txn_type, keywords in TYPE_KEYWORDS:
        if any(kw in desc_lower for kw in keywords):
            return txn_type
    if paid_in > 0 and withdrawn == 0.0:
        return "Money In"
    if withdrawn < 0 and paid_in == 0.0:
        return "Money Out"
    return "Other"


# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — DATA CLEANING & VALIDATION
# ─────────────────────────────────────────────────────────────────────────────
def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    print(f"[INFO] Rows before cleaning: {len(df)}")

    zero_mask = (
        (df["paid_in"] == 0.0)
        & (df["withdrawn"] == 0.0)
        & (df["balance"] == 0.0)
        & (df["status"] != "Completed")
    )
    df = df[~zero_mask].copy()

    df.drop_duplicates(
        subset=["receipt_no", "datetime", "details"],
        keep="first",
        inplace=True,
    )

    for col in ["paid_in", "withdrawn", "balance"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["datetime"] = pd.to_datetime(df["datetime"], errors="coerce")
    df.sort_values("datetime", ascending=True, inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.insert(0, "txn_no", range(1, len(df) + 1))

    print(f"[INFO] Rows after cleaning: {len(df)}")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — EXPORT
# ─────────────────────────────────────────────────────────────────────────────
def export_to_csv(df: pd.DataFrame, output_path: str) -> None:
    col_order = [
        "txn_no", "receipt_no", "date", "time", "details",
        "transaction_type", "status", "paid_in", "withdrawn", "balance",
    ]
    export_cols = [c for c in col_order if c in df.columns]
    df[export_cols].to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"\n[SUCCESS] CSV saved → {output_path}")
    print(f"          Rows : {len(df)}")
    print(f"          Cols : {export_cols}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    if not PDF_FOLDER.exists():
        sys.exit(f"[FATAL] Folder not found: '{PDF_FOLDER.resolve()}'\n"
                 "        Update PDF_FOLDER at the top of this script.")

    pdf_files = sorted(PDF_FOLDER.glob("*.pdf"))

    if not pdf_files:
        sys.exit(f"[FATAL] No PDF files found in '{PDF_FOLDER.resolve()}'")

    print("=" * 60)
    print("  M-PESA PDF STATEMENT EXTRACTOR — PART 1  (multi-file)")
    print("=" * 60)
    print(f"[INFO] PDF folder      : {PDF_FOLDER.resolve()}")
    print(f"[INFO] PDF files found : {len(pdf_files)}")
    for i, p in enumerate(pdf_files, 1):
        print(f"         {i:>3}. {p.name}")
    print()

    all_transactions = []

    for idx, pdf_path in enumerate(pdf_files, 1):
        print(f"[FILE {idx}/{len(pdf_files)}] Processing: {pdf_path.name}")
        pages = extract_raw_text(str(pdf_path))
        file_lines = []
        for page_text in pages:
            file_lines.extend(clean_page_lines(page_text))
        print(f"         Clean lines        : {len(file_lines)}")
        txns = parse_lines_into_transactions(file_lines)
        print(f"         Transactions found : {len(txns)}")
        all_transactions.extend(txns)
        print()

    if not all_transactions:
        sys.exit("[FATAL] No transactions parsed from any PDF.\n"
                 "        Confirm the PDFs have a text layer (not image-only).")

    print(f"[INFO] Total raw transactions across all files: {len(all_transactions)}")

    df = pd.DataFrame(all_transactions)

    # receipt_no is globally unique per M-PESA transaction — safe dedup key across overlapping statements
    rows_before = len(df)
    df["datetime"] = pd.to_datetime(df["datetime"], errors="coerce")
    df.sort_values(["receipt_no", "datetime"], inplace=True)
    df.drop_duplicates(subset=["receipt_no"], keep="first", inplace=True)
    print(f"[INFO] Duplicates removed         : {rows_before - len(df)}")
    print(f"[INFO] Unique transactions remain  : {len(df)}\n")

    df = clean_dataframe(df)

    print("\n── Column dtypes ────────────────────────────────────────")
    print(df[["receipt_no", "date", "time", "paid_in", "withdrawn", "balance"]].dtypes.to_string())

    print("\n── First 10 rows ────────────────────────────────────────")
    print(df[["receipt_no", "date", "time", "paid_in", "withdrawn", "balance", "details"]].head(10).to_string())

    print("\n── Financial Totals ─────────────────────────────────────")
    print(f"  Total Paid In    : KES {df['paid_in'].sum():>14,.2f}")
    print(f"  Total Withdrawn  : KES {df['withdrawn'].sum():>14,.2f}")
    print(f"  Net Cash Flow    : KES {(df['paid_in'].sum() + df['withdrawn'].sum()):>14,.2f}")
    print(f"  Date Range       : {df['date'].min()} → {df['date'].max()}")
    print(f"  Transaction Types:")
    print(df["transaction_type"].value_counts().to_string())
    print("─" * 55)

    output_csv = PDF_FOLDER / "MPESA_all_statements_clean.csv"
    export_to_csv(df, str(output_csv))
    display(df)


if __name__ == "__main__":
    main()
