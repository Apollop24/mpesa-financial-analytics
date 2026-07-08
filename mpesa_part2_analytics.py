# =============================================================================
# PART 2 — M-PESA FINANCIAL ANALYTICS ENGINE & INTERACTIVE DASHBOARD GENERATOR
# =============================================================================
# Continues directly from Part 1. Call run_analytics(df, output_dir) after
# clean_dataframe() returns df in main().
#
# Requirements:
#   pip install pandas numpy plotly openpyxl scipy scikit-learn --break-system-packages
#
# Outputs (all written to PDF_FOLDER / "mpesa_output"):
#   mpesa_clean.csv                — cleaned transaction dataset
#   mpesa_clean.xlsx               — Excel workbook with multiple sheets
#   mpesa_dashboard.html           — self-contained interactive dashboard
# =============================================================================

import re
import json
import warnings
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
from collections import Counter

warnings.filterwarnings("ignore")


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — FEATURE ENGINEERING
# ─────────────────────────────────────────────────────────────────────────────

def engineer_features(df: pd.DataFrame) -> pd.DataFrame:
    """Add time, behavioural and financial derived columns."""
    df = df.copy()

    df["datetime"] = pd.to_datetime(df["datetime"], errors="coerce")
    df["date"]     = pd.to_datetime(df["date"],     errors="coerce")

    # Time decomposition
    df["year"]        = df["datetime"].dt.year
    df["quarter"]     = df["datetime"].dt.quarter
    df["month"]       = df["datetime"].dt.month
    df["month_name"]  = df["datetime"].dt.strftime("%B")
    df["week"]        = df["datetime"].dt.isocalendar().week.astype(int)
    df["day"]         = df["datetime"].dt.day
    df["day_name"]    = df["datetime"].dt.strftime("%A")
    df["hour"]        = df["datetime"].dt.hour
    df["minute"]      = df["datetime"].dt.minute
    df["is_weekend"]  = df["datetime"].dt.dayofweek >= 5
    df["is_business_day"] = ~df["is_weekend"]

    # Time-of-day buckets
    def time_bucket(h):
        if 5 <= h < 12:  return "Morning"
        if 12 <= h < 17: return "Afternoon"
        if 17 <= h < 21: return "Evening"
        return "Night"
    df["time_of_day"] = df["hour"].apply(time_bucket)

    # Financial flags
    df["is_income"]  = df["paid_in"] > 0
    df["is_expense"] = df["withdrawn"] < 0
    df["amount"]     = df["paid_in"] + df["withdrawn"]          # signed net per row
    df["abs_amount"] = df["amount"].abs()

    # Cumulative columns (sorted chronologically)
    df = df.sort_values("datetime").reset_index(drop=True)
    df["cumulative_income"]  = df["paid_in"].cumsum()
    df["cumulative_expense"] = df["withdrawn"].abs().cumsum()
    df["cumulative_savings"] = df["paid_in"].cumsum() + df["withdrawn"].cumsum()

    # Month-year label for grouping
    df["month_year"] = df["datetime"].dt.strftime("%b %Y")
    df["week_year"]  = df["datetime"].dt.strftime("W%W %Y")

    return df


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — ENTITY EXTRACTION (WHO SENT / WHO RECEIVED)
# ─────────────────────────────────────────────────────────────────────────────

SEND_KEYWORDS  = ["transfer to", "send money", "customer transfer to",
                  "payment to", "merchant payment to", "pay bill to",
                  "customer payment to", "sent to"]
RECV_KEYWORDS  = ["funds received from", "received from", "transfer from",
                  "business payment from", "salary from", "payment from"]

def _extract_entity(detail: str, keywords: list) -> str:
    """Pull the entity name/number that follows a keyword in a detail string."""
    detail_lower = detail.lower()
    for kw in sorted(keywords, key=len, reverse=True):
        if kw in detail_lower:
            idx = detail_lower.index(kw) + len(kw)
            raw = detail[idx:].strip(" -–—:").split("  ")[0].strip()
            # Trim trailing noise (amounts, receipt codes, status words)
            raw = re.sub(r"\s+(Completed|Failed|Reversed|KES|KSH).*$", "", raw, flags=re.IGNORECASE)
            return raw.strip() if raw else "Unknown"
    return None


def extract_entities(df: pd.DataFrame):
    """
    Returns two DataFrames:
      top_recipients — people/merchants/paybills I sent money to
      top_senders    — people/merchants/employers who sent money to me
    """
    recipients = []
    senders    = []

    for _, row in df.iterrows():
        detail = str(row.get("details", ""))
        amt    = abs(row.get("withdrawn", 0))
        inc    = row.get("paid_in", 0)

        if amt > 0:
            entity = _extract_entity(detail, SEND_KEYWORDS)
            if not entity:
                entity = detail.strip()
            recipients.append({"entity": entity, "amount": amt,
                                "txn_type": row.get("transaction_type",""),
                                "detail": detail})

        if inc > 0:
            entity = _extract_entity(detail, RECV_KEYWORDS)
            if not entity:
                entity = detail.strip()
            senders.append({"entity": entity, "amount": inc,
                            "txn_type": row.get("transaction_type",""),
                            "detail": detail})

    def _summarise(rows):
        if not rows:
            return pd.DataFrame(columns=["entity","total_amount","txn_count","avg_amount"])
        tmp = pd.DataFrame(rows)
        grp = tmp.groupby("entity")["amount"].agg(
            total_amount="sum", txn_count="count", avg_amount="mean"
        ).reset_index().sort_values("total_amount", ascending=False).head(30)
        grp[["total_amount","avg_amount"]] = grp[["total_amount","avg_amount"]].round(2)
        return grp

    return _summarise(recipients), _summarise(senders)


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — CORE FINANCIAL METRICS
# ─────────────────────────────────────────────────────────────────────────────

def compute_metrics(df: pd.DataFrame) -> dict:
    income_rows  = df[df["is_income"]]
    expense_rows = df[df["is_expense"]]

    total_in   = df["paid_in"].sum()
    total_out  = df["withdrawn"].abs().sum()
    net_flow   = total_in - total_out

    date_range_days = max((df["datetime"].max() - df["datetime"].min()).days, 1)

    # ── Income metrics ──────────────────────────────────────────────────────
    income_vals = income_rows["paid_in"]
    income = dict(
        total          = round(total_in, 2),
        count          = len(income_rows),
        mean           = round(income_vals.mean(), 2) if len(income_vals) else 0,
        median         = round(income_vals.median(), 2) if len(income_vals) else 0,
        std            = round(income_vals.std(), 2) if len(income_vals) else 0,
        max            = round(income_vals.max(), 2) if len(income_vals) else 0,
        min            = round(income_vals.min(), 2) if len(income_vals) else 0,
        daily_avg      = round(total_in / date_range_days, 2),
        weekly_avg     = round(total_in / (date_range_days / 7), 2),
        monthly_avg    = round(total_in / max(df["month_year"].nunique(), 1), 2),
        volatility_cv  = round((income_vals.std() / income_vals.mean() * 100), 2) if income_vals.mean() else 0,
    )

    # ── Expense metrics ─────────────────────────────────────────────────────
    expense_vals = expense_rows["withdrawn"].abs()
    expense = dict(
        total          = round(total_out, 2),
        count          = len(expense_rows),
        mean           = round(expense_vals.mean(), 2) if len(expense_vals) else 0,
        median         = round(expense_vals.median(), 2) if len(expense_vals) else 0,
        std            = round(expense_vals.std(), 2) if len(expense_vals) else 0,
        max            = round(expense_vals.max(), 2) if len(expense_vals) else 0,
        min            = round(expense_vals.min(), 2) if len(expense_vals) else 0,
        daily_avg      = round(total_out / date_range_days, 2),
        weekly_avg     = round(total_out / (date_range_days / 7), 2),
        monthly_avg    = round(total_out / max(df["month_year"].nunique(), 1), 2),
        volatility_cv  = round((expense_vals.std() / expense_vals.mean() * 100), 2) if expense_vals.mean() and len(expense_vals) else 0,
    )

    # ── Savings ─────────────────────────────────────────────────────────────
    savings_rate = round((net_flow / total_in * 100), 2) if total_in else 0
    savings = dict(
        net_flow     = round(net_flow, 2),
        savings_rate = savings_rate,
        ratio        = round(net_flow / total_in, 4) if total_in else 0,
    )

    # ── Balance ─────────────────────────────────────────────────────────────
    bal = df["balance"]
    balance = dict(
        latest  = round(bal.iloc[-1], 2) if len(bal) else 0,
        highest = round(bal.max(), 2)    if len(bal) else 0,
        lowest  = round(bal.min(), 2)    if len(bal) else 0,
        mean    = round(bal.mean(), 2)   if len(bal) else 0,
        std     = round(bal.std(), 2)    if len(bal) else 0,
        volatility_cv = round(bal.std() / bal.mean() * 100, 2) if bal.mean() else 0,
    )

    # ── Health score (0-100) ─────────────────────────────────────────────────
    score = 0
    score += min(savings_rate, 30)                              # savings weight
    score += min(30, 30 * (total_in / max(total_out, 1) - 1))  # income > expense
    score += min(20, 20 * (1 - expense["volatility_cv"] / 200))
    score += min(20, 20 * (1 - income["volatility_cv"] / 200))
    health_score = max(0, min(100, round(score, 1)))

    # ── Monthly cash flow ────────────────────────────────────────────────────
    monthly_cf = df.groupby("month_year").agg(
        paid_in   = ("paid_in",   "sum"),
        withdrawn = ("withdrawn", lambda x: x.abs().sum()),
    ).reset_index()
    monthly_cf["net"] = monthly_cf["paid_in"] - monthly_cf["withdrawn"]
    monthly_cf = monthly_cf.sort_values("month_year")

    best_month  = monthly_cf.loc[monthly_cf["net"].idxmax(), "month_year"]  if len(monthly_cf) else "N/A"
    worst_month = monthly_cf.loc[monthly_cf["net"].idxmin(), "month_year"] if len(monthly_cf) else "N/A"

    # ── Charges ──────────────────────────────────────────────────────────────
    charges_total = df[df["transaction_type"] == "Charges / Fees"]["withdrawn"].abs().sum()
    reversed_count = (df["status"] == "Reversed").sum()

    return dict(
        income=income, expense=expense, savings=savings, balance=balance,
        health_score=health_score, best_month=best_month, worst_month=worst_month,
        total_txns=len(df), date_range_days=date_range_days,
        charges_total=round(charges_total, 2), reversed_count=int(reversed_count),
        monthly_cf=monthly_cf,
    )


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — ANOMALY DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def detect_anomalies(df: pd.DataFrame) -> pd.DataFrame:
    """Flag transactions more than 2 std deviations from the mean (Z-score)."""
    df = df.copy()
    df["z_score"] = 0.0
    df["is_anomaly"] = False

    amounts = df["abs_amount"]
    if len(amounts) > 3:
        mu, sigma = amounts.mean(), amounts.std()
        if sigma > 0:
            df["z_score"]   = ((amounts - mu) / sigma).round(2)
            df["is_anomaly"] = df["z_score"].abs() > 2.5

    return df


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────

def export_excel(df: pd.DataFrame, metrics: dict,
                 top_recipients: pd.DataFrame, top_senders: pd.DataFrame,
                 path: Path) -> None:
    """Write multi-sheet Excel workbook."""
    try:
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("[WARN] openpyxl not installed — skipping Excel export.")
        return

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # Sheet 1 — transactions
        export_cols = [c for c in ["txn_no","receipt_no","date","time","details",
                                   "transaction_type","status","paid_in",
                                   "withdrawn","balance","z_score","is_anomaly"]
                       if c in df.columns]
        df[export_cols].to_excel(writer, sheet_name="Transactions", index=False)

        # Sheet 2 — monthly summary
        monthly = df.groupby(["year","month","month_name"], as_index=False).agg(
            total_in   = ("paid_in",   "sum"),
            total_out  = ("withdrawn", lambda x: x.abs().sum()),
            txn_count  = ("paid_in",   "count"),
        )
        monthly["net_flow"]     = monthly["total_in"] - monthly["total_out"]
        monthly["savings_rate"] = (monthly["net_flow"] / monthly["total_in"].replace(0, np.nan) * 100).round(2)
        monthly.to_excel(writer, sheet_name="Monthly Summary", index=False)

        # Sheet 3 — by category
        by_type = df.groupby("transaction_type", as_index=False).agg(
            count      = ("txn_no",   "count"),
            total_in   = ("paid_in",   "sum"),
            total_out  = ("withdrawn", lambda x: x.abs().sum()),
        )
        by_type["net"] = by_type["total_in"] - by_type["total_out"]
        by_type.to_excel(writer, sheet_name="By Category", index=False)

        # Sheet 4 — top recipients
        top_recipients.to_excel(writer, sheet_name="Top Recipients", index=False)

        # Sheet 5 — top senders
        top_senders.to_excel(writer, sheet_name="Top Senders", index=False)

        # Sheet 6 — anomalies
        anomalies = df[df["is_anomaly"]][export_cols]
        anomalies.to_excel(writer, sheet_name="Anomalies", index=False)

    print(f"[SUCCESS] Excel saved → {path}")


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 — HTML DASHBOARD (self-contained, data injected as JSON)
# ─────────────────────────────────────────────────────────────────────────────

def _safe_json(obj):
    """Convert numpy/pandas types to plain Python for json.dumps."""
    if isinstance(obj, (np.integer,)):  return int(obj)
    if isinstance(obj, (np.floating,)): return float(obj)
    if isinstance(obj, (np.bool_,)):    return bool(obj)
    if isinstance(obj, pd.Timestamp):   return str(obj)
    if isinstance(obj, float) and np.isnan(obj): return 0
    return obj


def _df_to_records(df: pd.DataFrame) -> list:
    """Convert a dataframe to a list of plain-python dicts safe for JSON."""
    records = []
    for row in df.to_dict(orient="records"):
        records.append({k: _safe_json(v) for k, v in row.items()})
    return records


def build_dashboard(df: pd.DataFrame, metrics: dict,
                    top_recipients: pd.DataFrame, top_senders: pd.DataFrame,
                    output_path: Path) -> None:
    """Generate self-contained HTML dashboard with all data embedded as JSON."""

    # ── Prepare JSON data bundles ──────────────────────────────────────────
    monthly_cf = metrics["monthly_cf"]

    by_type = df.groupby("transaction_type", as_index=False).agg(
        count    = ("txn_no",   "count"),
        total_in = ("paid_in",   "sum"),
        total_out= ("withdrawn", lambda x: round(x.abs().sum(), 2)),
    ).sort_values("total_out", ascending=False)
    by_type["net"] = (by_type["total_in"] - by_type["total_out"]).round(2)

    by_hour = df.groupby("hour").agg(
        count    = ("txn_no",   "count"),
        total_out= ("withdrawn", lambda x: x.abs().sum()),
    ).reset_index()

    by_dow = df.groupby("day_name").agg(
        count    = ("txn_no",   "count"),
        total_out= ("withdrawn", lambda x: x.abs().sum()),
    ).reset_index()
    dow_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    by_dow["sort_key"] = by_dow["day_name"].map({d:i for i,d in enumerate(dow_order)})
    by_dow = by_dow.sort_values("sort_key")

    # Balance timeline
    bal_timeline = df[["date","balance","paid_in","withdrawn","details"]].copy()
    bal_timeline["date"] = bal_timeline["date"].astype(str)

    # All transactions for the table (most recent first)
    all_txns = df.sort_values("datetime", ascending=False)[
        ["txn_no","receipt_no","date","time","details",
         "transaction_type","status","paid_in","withdrawn","balance","is_anomaly"]
    ].copy()
    all_txns["date"] = all_txns["date"].astype(str)
    all_txns["is_anomaly"] = all_txns["is_anomaly"].astype(bool)

    # Anomalies
    anomalies = df[df["is_anomaly"]].sort_values("abs_amount", ascending=False)[
        ["date","time","details","transaction_type","paid_in","withdrawn","balance","z_score"]
    ].copy()
    anomalies["date"] = anomalies["date"].astype(str)

    DATA = {
        "metrics":         metrics | {"monthly_cf": _df_to_records(monthly_cf)},
        "by_type":         _df_to_records(by_type),
        "by_hour":         _df_to_records(by_hour),
        "by_dow":          _df_to_records(by_dow),
        "bal_timeline":    _df_to_records(bal_timeline),
        "top_recipients":  _df_to_records(top_recipients),
        "top_senders":     _df_to_records(top_senders),
        "all_txns":        _df_to_records(all_txns),
        "anomalies":       _df_to_records(anomalies),
        "generated_at":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    # Remove non-serialisable nested objects before dump
    DATA["metrics"].pop("monthly_cf", None)
    DATA["metrics"]["monthly_cf"] = _df_to_records(monthly_cf)

    json_blob = json.dumps(DATA, default=_safe_json)

    # ── HTML template ──────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>M-PESA Financial Analytics Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
:root{{
  --bg:#0d1117;--surface:#161b22;--card:#1c2128;--border:#30363d;
  --text:#e6edf3;--muted:#8b949e;--accent:#2a78d6;--green:#1baf7a;
  --red:#e34948;--amber:#eda100;--purple:#7c6fcd;
  --font:'Segoe UI',system-ui,sans-serif;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:var(--font);font-size:14px;line-height:1.5}}
a{{color:var(--accent)}}
.shell{{display:flex;min-height:100vh}}
.sidebar{{width:220px;background:var(--surface);border-right:1px solid var(--border);
  padding:1.5rem 1rem;position:fixed;top:0;left:0;height:100vh;overflow-y:auto;z-index:100}}
.sidebar h1{{font-size:15px;font-weight:600;color:var(--text);margin-bottom:4px}}
.sidebar p{{font-size:11px;color:var(--muted);margin-bottom:1.5rem}}
.nav-item{{display:block;padding:8px 12px;border-radius:6px;cursor:pointer;
  font-size:13px;color:var(--muted);margin-bottom:2px;border:none;
  background:transparent;width:100%;text-align:left;transition:background 0.15s}}
.nav-item:hover{{background:var(--card);color:var(--text)}}
.nav-item.active{{background:var(--accent);color:#fff;font-weight:500}}
.main{{margin-left:220px;padding:1.5rem;width:calc(100% - 220px)}}
.page{{display:none}}.page.active{{display:block}}
.page-title{{font-size:20px;font-weight:600;margin-bottom:6px}}
.page-sub{{font-size:13px;color:var(--muted);margin-bottom:1.5rem}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:1.5rem}}
.kpi{{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:16px}}
.kpi-label{{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}}
.kpi-value{{font-size:22px;font-weight:600;color:var(--text)}}
.kpi-delta{{font-size:11px;margin-top:4px}}
.green{{color:var(--green)}}.red{{color:var(--red)}}.amber{{color:var(--amber)}}.muted{{color:var(--muted)}}
.grid-2{{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:14px;margin-bottom:1.5rem}}
.grid-3{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-bottom:1.5rem}}
.card{{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:16px}}
.card-title{{font-size:13px;font-weight:500;color:var(--text);margin-bottom:12px}}
.chart-wrap{{position:relative;width:100%}}
.filters{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:1.2rem;align-items:flex-end}}
.filters label{{font-size:11px;color:var(--muted);display:block;margin-bottom:4px}}
.filters select,.filters input{{
  background:var(--card);border:1px solid var(--border);color:var(--text);
  padding:7px 10px;border-radius:6px;font-size:13px;min-width:130px}}
.filters button{{padding:8px 14px;border-radius:6px;border:1px solid var(--border);
  background:var(--accent);color:#fff;font-size:13px;cursor:pointer}}
.tbl-wrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{background:var(--surface);color:var(--muted);padding:9px 10px;text-align:left;
  border-bottom:1px solid var(--border);white-space:nowrap;font-weight:500}}
td{{padding:8px 10px;border-bottom:1px solid var(--border);color:var(--text);vertical-align:middle}}
tr:hover td{{background:rgba(255,255,255,0.02)}}
.badge{{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:500}}
.badge-in{{background:rgba(27,175,122,.15);color:#1baf7a}}
.badge-out{{background:rgba(227,73,72,.15);color:#e34948}}
.badge-ok{{background:rgba(27,175,122,.15);color:#1baf7a}}
.badge-rev{{background:rgba(237,161,0,.15);color:#eda100}}
.badge-fail{{background:rgba(227,73,72,.15);color:#e34948}}
.badge-anom{{background:rgba(124,111,205,.2);color:#7c6fcd}}
.health-ring{{display:flex;align-items:center;gap:20px;padding:12px 0}}
.score-num{{font-size:52px;font-weight:700}}
.score-label{{font-size:13px;color:var(--muted)}}
.finding{{border-left:3px solid var(--accent);padding:10px 14px;margin-bottom:10px;
  background:rgba(255,255,255,0.02);border-radius:0 6px 6px 0}}
.finding.good{{border-color:var(--green)}}.finding.warn{{border-color:var(--amber)}}.finding.danger{{border-color:var(--red)}}
.finding-title{{font-size:12px;font-weight:500;margin-bottom:3px}}
.finding-val{{font-size:18px;font-weight:600}}
.finding-desc{{font-size:11px;color:var(--muted);margin-top:3px}}
.pagination{{display:flex;gap:8px;align-items:center;justify-content:flex-end;margin-top:12px}}
.pagination button{{padding:5px 12px;border-radius:5px;border:1px solid var(--border);
  background:var(--card);color:var(--text);cursor:pointer;font-size:12px}}
.pagination span{{font-size:12px;color:var(--muted)}}
.gen-info{{font-size:11px;color:var(--muted);margin-top:2rem;text-align:center}}
</style>
</head>
<body>
<div class="shell">
<nav class="sidebar">
  <h1>M-PESA Analytics</h1>
  <p id="genAt"></p>
  <button class="nav-item active" onclick="nav('overview',this)">&#9632; Overview</button>
  <button class="nav-item" onclick="nav('cashflow',this)">&#9196; Cash Flow</button>
  <button class="nav-item" onclick="nav('income',this)">&#8679; Income</button>
  <button class="nav-item" onclick="nav('spending',this)">&#8681; Spending</button>
  <button class="nav-item" onclick="nav('behavior',this)">&#9200; Behavior</button>
  <button class="nav-item" onclick="nav('entities',this)">&#128101; Who Sends / Receives</button>
  <button class="nav-item" onclick="nav('health',this)">&#10084; Financial Health</button>
  <button class="nav-item" onclick="nav('anomalies',this)">&#9888; Anomalies</button>
  <button class="nav-item" onclick="nav('transactions',this)">&#9776; All Transactions</button>
</nav>

<main class="main">

<!-- ═══════ OVERVIEW ═══════ -->
<div id="page-overview" class="page active">
  <div class="page-title">Executive overview</div>
  <div class="page-sub" id="overviewSub"></div>
  <div class="kpi-grid" id="kpiRow"></div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">Monthly income vs spending (KES)</div>
      <div class="chart-wrap" style="height:240px"><canvas id="c-monthly" role="img" aria-label="Monthly income vs spending">Monthly chart</canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Spending by category</div>
      <div class="chart-wrap" style="height:240px"><canvas id="c-donut" role="img" aria-label="Spending by category">Category donut</canvas></div>
    </div>
  </div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">Running balance (KES)</div>
      <div class="chart-wrap" style="height:220px"><canvas id="c-balance" role="img" aria-label="Running balance over time">Balance chart</canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Transaction status breakdown</div>
      <div class="chart-wrap" style="height:220px"><canvas id="c-status" role="img" aria-label="Status breakdown">Status pie</canvas></div>
    </div>
  </div>
</div>

<!-- ═══════ CASH FLOW ═══════ -->
<div id="page-cashflow" class="page">
  <div class="page-title">Cash flow analysis</div>
  <div class="page-sub">Monthly net surplus / deficit and cumulative savings trajectory</div>
  <div class="kpi-grid" id="cfKpi"></div>
  <div class="card" style="margin-bottom:1.2rem">
    <div class="card-title">Monthly net cash flow (KES)</div>
    <div class="chart-wrap" style="height:260px"><canvas id="c-netflow" role="img" aria-label="Net cash flow by month">Net flow chart</canvas></div>
  </div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">Cumulative savings trajectory (KES)</div>
      <div class="chart-wrap" style="height:220px"><canvas id="c-cumsav" role="img" aria-label="Cumulative savings">Cumulative savings</canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Monthly savings rate (%)</div>
      <div class="chart-wrap" style="height:220px"><canvas id="c-savrate" role="img" aria-label="Monthly savings rate">Savings rate</canvas></div>
    </div>
  </div>
</div>

<!-- ═══════ INCOME ═══════ -->
<div id="page-income" class="page">
  <div class="page-title">Income analytics</div>
  <div class="page-sub">All money received — by source, time and distribution</div>
  <div class="kpi-grid" id="incKpi"></div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">Monthly income trend (KES)</div>
      <div class="chart-wrap" style="height:230px"><canvas id="c-inctrend" role="img" aria-label="Monthly income trend">Income trend</canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Income by category</div>
      <div class="chart-wrap" style="height:230px"><canvas id="c-inccat" role="img" aria-label="Income by category">Income category</canvas></div>
    </div>
  </div>
</div>

<!-- ═══════ SPENDING ═══════ -->
<div id="page-spending" class="page">
  <div class="page-title">Spending analytics</div>
  <div class="page-sub">Outflow breakdown by category, trend and distribution</div>
  <div class="kpi-grid" id="spKpi"></div>
  <div class="card" style="margin-bottom:1.2rem">
    <div class="card-title">Total spending by category (KES)</div>
    <div class="chart-wrap" id="c-spcat-wrap" style="height:360px"><canvas id="c-spcat" role="img" aria-label="Spending by category horizontal bar">Spending categories</canvas></div>
  </div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">Monthly spending trend (KES)</div>
      <div class="chart-wrap" style="height:220px"><canvas id="c-sptrend" role="img" aria-label="Monthly spending trend">Spending trend</canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Spending distribution (KES)</div>
      <div class="chart-wrap" style="height:220px"><canvas id="c-sphist" role="img" aria-label="Spending histogram">Histogram</canvas></div>
    </div>
  </div>
</div>

<!-- ═══════ BEHAVIOR ═══════ -->
<div id="page-behavior" class="page">
  <div class="page-title">Behavioral analytics</div>
  <div class="page-sub">When do you transact? Activity patterns by hour and day</div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">Transaction count by hour of day</div>
      <div class="chart-wrap" style="height:230px"><canvas id="c-hour" role="img" aria-label="Hourly activity">Hourly activity</canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Transaction count by day of week</div>
      <div class="chart-wrap" style="height:230px"><canvas id="c-dow" role="img" aria-label="Day of week activity">Day of week</canvas></div>
    </div>
  </div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">Spending by hour of day (KES)</div>
      <div class="chart-wrap" style="height:220px"><canvas id="c-hour-sp" role="img" aria-label="Hourly spending">Hourly spending</canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Spending by day of week (KES)</div>
      <div class="chart-wrap" style="height:220px"><canvas id="c-dow-sp" role="img" aria-label="Day of week spending">DOW spending</canvas></div>
    </div>
  </div>
</div>

<!-- ═══════ ENTITIES ═══════ -->
<div id="page-entities" class="page">
  <div class="page-title">Who sends &amp; who receives</div>
  <div class="page-sub">Ranked by total KES amount — extracted from transaction details</div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">&#8679; Top recipients — people &amp; merchants I pay (KES sent)</div>
      <div class="chart-wrap" id="c-recip-wrap" style="height:380px"><canvas id="c-recip" role="img" aria-label="Top recipients">Recipients chart</canvas></div>
      <div class="tbl-wrap" style="margin-top:14px">
        <table>
          <thead><tr><th>#</th><th>Recipient / Merchant</th><th>Total sent (KES)</th><th>Transactions</th><th>Avg per txn (KES)</th></tr></thead>
          <tbody id="recipTbl"></tbody>
        </table>
      </div>
    </div>
    <div class="card">
      <div class="card-title">&#8681; Top senders — people &amp; sources who pay me (KES received)</div>
      <div class="chart-wrap" id="c-send-wrap" style="height:380px"><canvas id="c-send" role="img" aria-label="Top senders">Senders chart</canvas></div>
      <div class="tbl-wrap" style="margin-top:14px">
        <table>
          <thead><tr><th>#</th><th>Sender / Source</th><th>Total received (KES)</th><th>Transactions</th><th>Avg per txn (KES)</th></tr></thead>
          <tbody id="sendTbl"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ═══════ HEALTH ═══════ -->
<div id="page-health" class="page">
  <div class="page-title">Financial health indicators</div>
  <div class="page-sub">Composite score and key ratios derived from your statement</div>
  <div class="grid-2" style="margin-bottom:1.2rem">
    <div class="card">
      <div class="card-title">Financial health score</div>
      <div class="health-ring">
        <div>
          <div class="score-num" id="healthScore"></div>
          <div class="score-label">out of 100</div>
        </div>
        <div style="flex:1">
          <div class="chart-wrap" style="height:180px"><canvas id="c-radar" role="img" aria-label="Financial health radar">Radar chart</canvas></div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-title">Key financial ratios</div>
      <div id="ratioTable"></div>
    </div>
  </div>
  <div class="card">
    <div class="card-title">Automated insights</div>
    <div id="insightsList"></div>
  </div>
</div>

<!-- ═══════ ANOMALIES ═══════ -->
<div id="page-anomalies" class="page">
  <div class="page-title">Anomaly detection</div>
  <div class="page-sub">Transactions with Z-score &gt; 2.5 — statistically unusual amounts</div>
  <div class="kpi-grid" id="anomKpi"></div>
  <div class="card">
    <div class="card-title">Flagged transactions</div>
    <div class="tbl-wrap">
      <table>
        <thead><tr><th>Date</th><th>Time</th><th>Details</th><th>Type</th><th>Paid in</th><th>Withdrawn</th><th>Balance</th><th>Z-score</th></tr></thead>
        <tbody id="anomTbl"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ═══════ ALL TRANSACTIONS ═══════ -->
<div id="page-transactions" class="page">
  <div class="page-title">All transactions</div>
  <div class="page-sub">Full ledger — filter, search, paginate</div>
  <div class="filters">
    <div><label>Month</label>
      <select id="fMonth" onchange="filterTxns()">
        <option value="">All months</option>
      </select>
    </div>
    <div><label>Type</label>
      <select id="fType" onchange="filterTxns()">
        <option value="">All types</option>
      </select>
    </div>
    <div><label>Status</label>
      <select id="fStatus" onchange="filterTxns()">
        <option value="">All</option>
        <option>Completed</option><option>Failed</option><option>Reversed</option>
      </select>
    </div>
    <div><label>Search details</label>
      <input type="text" id="fSearch" placeholder="Name, merchant, ref…" oninput="filterTxns()">
    </div>
    <button onclick="resetTxnFilters()">Reset</button>
    <span id="txnCountLabel" style="font-size:12px;color:var(--muted);align-self:flex-end;padding-bottom:2px"></span>
  </div>
  <div class="card">
    <div class="tbl-wrap">
      <table>
        <thead><tr><th>#</th><th>Date</th><th>Time</th><th>Details</th><th>Type</th><th>Status</th><th style="text-align:right">Paid in</th><th style="text-align:right">Withdrawn</th><th style="text-align:right">Balance</th><th>Flag</th></tr></thead>
        <tbody id="txnTblBody"></tbody>
      </table>
    </div>
    <div class="pagination">
      <button onclick="txnPage(-1)">&#8592; Prev</button>
      <span id="txnPageInfo"></span>
      <button onclick="txnPage(1)">Next &#8594;</button>
    </div>
  </div>
</div>

<div class="gen-info" id="genInfo"></div>
</main>
</div>

<script>
const RAW = {json_blob};

// ── helpers ──────────────────────────────────────────────────────────────────
const fmt  = (n,d=2)=>Math.abs(n).toLocaleString('en-KE',{{minimumFractionDigits:d,maximumFractionDigits:d}});
const fmtS = (n)=>n.toLocaleString('en-KE',{{minimumFractionDigits:2,maximumFractionDigits:2}});
const COLORS=['#2a78d6','#1baf7a','#eda100','#7c6fcd','#e34948','#e87ba4','#eb6834','#008300'];
const charts={{}};
function mkChart(id,cfg){{if(charts[id])charts[id].destroy();charts[id]=new Chart(document.getElementById(id),cfg);}}

// ── nav ───────────────────────────────────────────────────────────────────────
function nav(page,btn){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  document.getElementById('page-'+page).classList.add('active');
  btn.classList.add('active');
  renderPage(page);
}}

// ── page router ───────────────────────────────────────────────────────────────
const rendered={{}};
function renderPage(page){{
  if(rendered[page])return;
  rendered[page]=true;
  if(page==='overview')   renderOverview();
  if(page==='cashflow')   renderCashflow();
  if(page==='income')     renderIncome();
  if(page==='spending')   renderSpending();
  if(page==='behavior')   renderBehavior();
  if(page==='entities')   renderEntities();
  if(page==='health')     renderHealth();
  if(page==='anomalies')  renderAnomalies();
  if(page==='transactions'){{ initTxnFilters(); filterTxns(); }}
}}

// ── KPI helper ────────────────────────────────────────────────────────────────
function kpiCard(label,val,delta,cls){{
  return `<div class="kpi"><div class="kpi-label">${{label}}</div>
    <div class="kpi-value">${{val}}</div>
    ${{delta?`<div class="kpi-delta ${{cls}}">${{delta}}</div>`:''}}
  </div>`;
}}

// ── OVERVIEW ──────────────────────────────────────────────────────────────────
function renderOverview(){{
  const m=RAW.metrics;
  const inc=m.income; const exp=m.expense; const sav=m.savings; const bal=m.balance;

  document.getElementById('overviewSub').textContent=
    `${{m.total_txns}} transactions · ${{m.date_range_days}} days · generated ${{RAW.generated_at}}`;

  document.getElementById('kpiRow').innerHTML=[
    kpiCard('Total paid in','KES '+fmt(inc.total),null,''),
    kpiCard('Total withdrawn','KES '+fmt(exp.total),null,''),
    kpiCard('Net cash flow',(sav.net_flow>=0?'+':'-')+' KES '+fmt(Math.abs(sav.net_flow)),
      sav.net_flow>=0?'Surplus':'Deficit',sav.net_flow>=0?'green':'red'),
    kpiCard('Savings rate',sav.savings_rate.toFixed(1)+'%',
      sav.savings_rate>20?'Healthy':sav.savings_rate>0?'Low':'Negative',
      sav.savings_rate>20?'green':sav.savings_rate>0?'amber':'red'),
    kpiCard('Latest balance','KES '+fmt(bal.latest),null,''),
    kpiCard('Total transactions',m.total_txns.toLocaleString(),null,''),
    kpiCard('Health score',m.health_score+'/100',null,''),
    kpiCard('Fees paid','KES '+fmt(m.charges_total),null,'red'),
  ].join('');

  // Monthly income vs spending
  const mc=m.monthly_cf;
  const mLabels=mc.map(r=>r.month_year);
  mkChart('c-monthly',{{
    type:'bar',
    data:{{labels:mLabels,datasets:[
      {{label:'Paid in',data:mc.map(r=>r.paid_in),backgroundColor:'#1baf7a'}},
      {{label:'Withdrawn',data:mc.map(r=>r.withdrawn),backgroundColor:'#e34948'}},
    ]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{labels:{{color:'#8b949e',font:{{size:11}}}}}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e',maxRotation:45,font:{{size:10}}}}}}}}}}
  }});

  // Donut — spending by type
  const bt=RAW.by_type.filter(r=>r.total_out>0).slice(0,8);
  mkChart('c-donut',{{
    type:'doughnut',
    data:{{labels:bt.map(r=>r.transaction_type),
           datasets:[{{data:bt.map(r=>r.total_out),backgroundColor:COLORS,borderWidth:2,borderColor:'#1c2128'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,cutout:'60%',
      plugins:{{legend:{{position:'right',labels:{{color:'#8b949e',font:{{size:11}},boxWidth:12}}}}}}}}
  }});

  // Balance line
  const bt2=RAW.bal_timeline;
  mkChart('c-balance',{{
    type:'line',
    data:{{labels:bt2.map(r=>r.date),
           datasets:[{{label:'Balance',data:bt2.map(r=>r.balance),
             borderColor:'#2a78d6',backgroundColor:'rgba(42,120,214,0.08)',
             borderWidth:2,fill:true,tension:0.3,pointRadius:0}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e',maxTicksLimit:8,autoSkip:true}}}}}}}}
  }});

  // Status pie
  const statusMap={{}};
  RAW.all_txns.forEach(r=>{{statusMap[r.status]=(statusMap[r.status]||0)+1;}});
  const sKeys=Object.keys(statusMap);
  const sCols=['#1baf7a','#e34948','#eda100'];
  mkChart('c-status',{{
    type:'pie',
    data:{{labels:sKeys,datasets:[{{data:sKeys.map(k=>statusMap[k]),backgroundColor:sCols,borderWidth:2,borderColor:'#1c2128'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'right',labels:{{color:'#8b949e',font:{{size:11}},boxWidth:12}}}}}}}}
  }});
}}

// ── CASH FLOW ─────────────────────────────────────────────────────────────────
function renderCashflow(){{
  const m=RAW.metrics; const mc=m.monthly_cf;
  const net=m.savings.net_flow;
  document.getElementById('cfKpi').innerHTML=[
    kpiCard('Net cash flow',(net>=0?'+':'-')+' KES '+fmt(Math.abs(net)),null,net>=0?'green':'red'),
    kpiCard('Best month',m.best_month,null,'green'),
    kpiCard('Worst month',m.worst_month,null,'red'),
    kpiCard('Savings rate',m.savings.savings_rate.toFixed(1)+'%',null,''),
  ].join('');

  const mLabels=mc.map(r=>r.month_year);
  const netVals=mc.map(r=>Math.round(r.net));

  mkChart('c-netflow',{{
    type:'bar',
    data:{{labels:mLabels,datasets:[{{
      label:'Net flow',data:netVals,
      backgroundColor:netVals.map(v=>v>=0?'rgba(27,175,122,0.75)':'rgba(227,73,72,0.75)')
    }}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e',maxRotation:45}}}}}}}}
  }});

  // Cumulative savings from all_txns
  let cumSum=0; const cumLabels=[]; const cumVals=[];
  [...RAW.all_txns].reverse().forEach(r=>{{
    cumSum+=(r.paid_in||0)-(r.withdrawn<0?Math.abs(r.withdrawn):0);
    cumLabels.push(r.date); cumVals.push(Math.round(cumSum));
  }});
  mkChart('c-cumsav',{{
    type:'line',
    data:{{labels:cumLabels,datasets:[{{label:'Cumulative savings',data:cumVals,
      borderColor:'#1baf7a',backgroundColor:'rgba(27,175,122,0.08)',
      borderWidth:2,fill:true,tension:0.3,pointRadius:0}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e',maxTicksLimit:8,autoSkip:true}}}}}}}}
  }});

  // Savings rate per month
  const srLabels=mc.map(r=>r.month_year);
  const srVals=mc.map(r=>r.paid_in>0?Math.round((r.net/r.paid_in)*100):0);
  mkChart('c-savrate',{{
    type:'line',
    data:{{labels:srLabels,datasets:[{{label:'Savings rate %',data:srVals,
      borderColor:'#eda100',backgroundColor:'rgba(237,161,0,0.08)',
      borderWidth:2,fill:true,tension:0.3,pointRadius:4,pointBackgroundColor:'#eda100'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>v+'%'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e',maxRotation:45}}}}}}}}
  }});
}}

// ── INCOME ────────────────────────────────────────────────────────────────────
function renderIncome(){{
  const inc=RAW.metrics.income;
  document.getElementById('incKpi').innerHTML=[
    kpiCard('Total income','KES '+fmt(inc.total),null,'green'),
    kpiCard('Transactions',inc.count,null,''),
    kpiCard('Mean income','KES '+fmt(inc.mean),null,''),
    kpiCard('Median income','KES '+fmt(inc.median),null,''),
    kpiCard('Largest','KES '+fmt(inc.max),null,''),
    kpiCard('Monthly avg','KES '+fmt(inc.monthly_avg),null,''),
    kpiCard('Volatility (CV)',inc.volatility_cv.toFixed(1)+'%',null,inc.volatility_cv>50?'amber':'green'),
  ].join('');

  const mc=RAW.metrics.monthly_cf;
  mkChart('c-inctrend',{{
    type:'line',
    data:{{labels:mc.map(r=>r.month_year),
           datasets:[{{label:'Paid in',data:mc.map(r=>Math.round(r.paid_in)),
             borderColor:'#1baf7a',backgroundColor:'rgba(27,175,122,0.08)',
             borderWidth:2,fill:true,tension:0.3,pointRadius:4,pointBackgroundColor:'#1baf7a'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e',maxRotation:45}}}}}}}}
  }});

  const incCat=RAW.by_type.filter(r=>r.total_in>0);
  mkChart('c-inccat',{{
    type:'doughnut',
    data:{{labels:incCat.map(r=>r.transaction_type),
           datasets:[{{data:incCat.map(r=>r.total_in),backgroundColor:COLORS,borderWidth:2,borderColor:'#1c2128'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,cutout:'55%',
      plugins:{{legend:{{position:'right',labels:{{color:'#8b949e',font:{{size:11}},boxWidth:12}}}}}}}}
  }});
}}

// ── SPENDING ──────────────────────────────────────────────────────────────────
function renderSpending(){{
  const exp=RAW.metrics.expense;
  document.getElementById('spKpi').innerHTML=[
    kpiCard('Total spending','KES '+fmt(exp.total),null,'red'),
    kpiCard('Transactions',exp.count,null,''),
    kpiCard('Mean expense','KES '+fmt(exp.mean),null,''),
    kpiCard('Median expense','KES '+fmt(exp.median),null,''),
    kpiCard('Largest','KES '+fmt(exp.max),null,''),
    kpiCard('Monthly avg','KES '+fmt(exp.monthly_avg),null,''),
    kpiCard('Volatility (CV)',exp.volatility_cv.toFixed(1)+'%',null,exp.volatility_cv>60?'amber':''),
  ].join('');

  const catData=RAW.by_type.filter(r=>r.total_out>0).sort((a,b)=>b.total_out-a.total_out);
  const catH=Math.max(320,catData.length*36+80);
  document.getElementById('c-spcat-wrap').style.height=catH+'px';
  mkChart('c-spcat',{{
    type:'bar',
    data:{{labels:catData.map(r=>r.transaction_type),
           datasets:[{{label:'Total spent',data:catData.map(r=>r.total_out),backgroundColor:'#e34948'}}]}},
    options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{x:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               y:{{ticks:{{color:'#8b949e'}}}}}}}}
  }});

  const mc=RAW.metrics.monthly_cf;
  mkChart('c-sptrend',{{
    type:'line',
    data:{{labels:mc.map(r=>r.month_year),
           datasets:[{{label:'Withdrawn',data:mc.map(r=>Math.round(r.withdrawn)),
             borderColor:'#e34948',backgroundColor:'rgba(227,73,72,0.08)',
             borderWidth:2,fill:true,tension:0.3,pointRadius:4,pointBackgroundColor:'#e34948'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e',maxRotation:45}}}}}}}}
  }});

  // Histogram from all_txns
  const outs=RAW.all_txns.filter(r=>r.withdrawn<0).map(r=>Math.abs(r.withdrawn));
  const bins=10; const maxV=Math.max(...outs); const binW=maxV/bins;
  const hist=Array(bins).fill(0);
  outs.forEach(v=>{{const i=Math.min(Math.floor(v/binW),bins-1);hist[i]++;}});
  const histLabels=Array.from({{length:bins}},(_,i)=>'KES '+Math.round(i*binW/1000)+'K');
  mkChart('c-sphist',{{
    type:'bar',
    data:{{labels:histLabels,
           datasets:[{{label:'Transactions',data:hist,backgroundColor:'#7c6fcd'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',stepSize:1}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e',maxRotation:45,font:{{size:10}}}}}}}}}}
  }});
}}

// ── BEHAVIOR ──────────────────────────────────────────────────────────────────
function renderBehavior(){{
  const bh=RAW.by_hour; const bd=RAW.by_dow;
  mkChart('c-hour',{{
    type:'bar',
    data:{{labels:bh.map(r=>r.hour+'h'),
           datasets:[{{label:'Transactions',data:bh.map(r=>r.count),backgroundColor:'#eda100'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',stepSize:1}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e'}}}}}}}}
  }});
  mkChart('c-dow',{{
    type:'bar',
    data:{{labels:bd.map(r=>r.day_name),
           datasets:[{{label:'Transactions',data:bd.map(r=>r.count),backgroundColor:'#2a78d6'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',stepSize:1}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e'}}}}}}}}
  }});
  mkChart('c-hour-sp',{{
    type:'bar',
    data:{{labels:bh.map(r=>r.hour+'h'),
           datasets:[{{label:'KES spent',data:bh.map(r=>Math.round(r.total_out||0)),backgroundColor:'#eb6834'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e'}}}}}}}}
  }});
  mkChart('c-dow-sp',{{
    type:'bar',
    data:{{labels:bd.map(r=>r.day_name),
           datasets:[{{label:'KES spent',data:bd.map(r=>Math.round(r.total_out||0)),backgroundColor:'#e87ba4'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{y:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
               x:{{ticks:{{color:'#8b949e'}}}}}}}}
  }});
}}

// ── ENTITIES ──────────────────────────────────────────────────────────────────
function renderEntities(){{
  function buildEntityChart(canvasId,wrapId,data,color,label){{
    const h=Math.max(320,data.length*34+80);
    document.getElementById(wrapId).style.height=h+'px';
    mkChart(canvasId,{{
      type:'bar',
      data:{{labels:data.map(r=>r.entity.length>30?r.entity.slice(0,30)+'…':r.entity),
             datasets:[{{label,data:data.map(r=>r.total_amount),backgroundColor:color}}]}},
      options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,
        plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{
          label:ctx=>'KES '+fmt(ctx.raw)+' ('+data[ctx.dataIndex].txn_count+' txns)'
        }}}}}},
        scales:{{x:{{ticks:{{color:'#8b949e',callback:v=>'KES '+Math.round(v/1000)+'K'}},grid:{{color:'rgba(255,255,255,0.05)'}}}},
                 y:{{ticks:{{color:'#8b949e',font:{{size:11}}}}}}}}}}
    }});
  }}

  buildEntityChart('c-recip','c-recip-wrap',RAW.top_recipients,'#e34948','KES sent');
  buildEntityChart('c-send','c-send-wrap',RAW.top_senders,'#1baf7a','KES received');

  document.getElementById('recipTbl').innerHTML=RAW.top_recipients.map((r,i)=>`
    <tr><td>${{i+1}}</td><td><strong>${{r.entity}}</strong></td>
    <td style="color:#e34948">KES ${{fmt(r.total_amount)}}</td>
    <td>${{r.txn_count}}</td><td>KES ${{fmt(r.avg_amount)}}</td></tr>`).join('');

  document.getElementById('sendTbl').innerHTML=RAW.top_senders.map((r,i)=>`
    <tr><td>${{i+1}}</td><td><strong>${{r.entity}}</strong></td>
    <td style="color:#1baf7a">KES ${{fmt(r.total_amount)}}</td>
    <td>${{r.txn_count}}</td><td>KES ${{fmt(r.avg_amount)}}</td></tr>`).join('');
}}

// ── HEALTH ────────────────────────────────────────────────────────────────────
function renderHealth(){{
  const m=RAW.metrics; const inc=m.income; const exp=m.expense;
  const sav=m.savings; const bal=m.balance;
  document.getElementById('healthScore').textContent=m.health_score;
  document.getElementById('healthScore').style.color=
    m.health_score>70?'#1baf7a':m.health_score>40?'#eda100':'#e34948';

  const ie_ratio=exp.total>0?(inc.total/exp.total).toFixed(2):'-';
  const ratios=[
    ['Income / expense ratio',ie_ratio,'> 1.0 is healthy',''],
    ['Savings rate',sav.savings_rate.toFixed(1)+'%','> 20% recommended',sav.savings_rate>20?'#1baf7a':'#eda100'],
    ['Balance volatility',bal.volatility_cv.toFixed(1)+'%','Lower = more stable',''],
    ['Income volatility',inc.volatility_cv.toFixed(1)+'%','Lower = more predictable',''],
    ['Expense volatility',exp.volatility_cv.toFixed(1)+'%','Lower = more controlled',''],
    ['Fees as % of income',(m.charges_total/inc.total*100).toFixed(2)+'%','Minimise where possible',''],
  ];
  document.getElementById('ratioTable').innerHTML=`<table style="width:100%">
    <thead><tr><th>Indicator</th><th>Value</th><th>Benchmark</th></tr></thead>
    <tbody>${{ratios.map(r=>`<tr>
      <td>${{r[0]}}</td>
      <td style="font-weight:500;color:${{r[3]||'var(--text)'}}">${{r[1]}}</td>
      <td style="color:var(--muted);font-size:11px">${{r[2]}}</td>
    </tr>`).join('')}}</tbody></table>`;

  // Radar — normalised 0-100 per dimension
  const radarLabels=['Savings rate','Income stability','Expense control','Balance health','Cash flow','Low fees'];
  const radarVals=[
    Math.min(100,sav.savings_rate*2),
    Math.max(0,100-inc.volatility_cv),
    Math.max(0,100-exp.volatility_cv),
    Math.min(100,(bal.lowest/Math.max(bal.mean,1))*100),
    Math.min(100,Math.max(0,(sav.net_flow/Math.max(inc.total,1))*100+50)),
    Math.max(0,100-(m.charges_total/Math.max(inc.total,1))*1000),
  ].map(v=>Math.round(Math.max(0,Math.min(100,v))));

  mkChart('c-radar',{{
    type:'radar',
    data:{{labels:radarLabels,datasets:[{{
      label:'Your profile',data:radarVals,
      backgroundColor:'rgba(42,120,214,0.15)',
      borderColor:'#2a78d6',pointBackgroundColor:'#2a78d6',borderWidth:2,pointRadius:4
    }}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{r:{{min:0,max:100,ticks:{{color:'#8b949e',backdropColor:'transparent',stepSize:20}},
               grid:{{color:'rgba(255,255,255,0.08)'}},
               pointLabels:{{color:'#8b949e',font:{{size:11}}}}}}}}}}
  }});

  // Insights
  const ins=[];
  if(sav.savings_rate>20)ins.push({{t:'Healthy savings rate',v:sav.savings_rate.toFixed(1)+'%',d:'You retain more than 20% of inflows — a strong financial habit.',c:'good'}});
  else if(sav.savings_rate>0)ins.push({{t:'Savings rate needs improvement',v:sav.savings_rate.toFixed(1)+'%',d:'Aim for at least 20% to build financial resilience.',c:'warn'}});
  else ins.push({{t:'Negative cash flow',v:sav.savings_rate.toFixed(1)+'%',d:'Outflows exceed inflows. Review recurring expenses.',c:'danger'}});

  if(m.charges_total>0)ins.push({{t:'Transaction fees paid',v:'KES '+fmt(m.charges_total),d:'Total M-PESA charges incurred during this period.',c:'warn'}});
  if(m.reversed_count>0)ins.push({{t:'Reversed transactions',v:m.reversed_count,d:'Check reversed transactions for potential issues.',c:'warn'}});

  const topExp=RAW.by_type.filter(r=>r.total_out>0).sort((a,b)=>b.total_out-a.total_out)[0];
  if(topExp)ins.push({{t:'Largest spending category',v:topExp.transaction_type,d:'KES '+fmt(topExp.total_out)+' — consider reviewing if this aligns with your priorities.',c:''}});

  const topRecip=RAW.top_recipients[0];
  if(topRecip)ins.push({{t:'Most paid recipient',v:topRecip.entity,d:'KES '+fmt(topRecip.total_amount)+' sent across '+topRecip.txn_count+' transactions.',c:''}});

  const topSend=RAW.top_senders[0];
  if(topSend)ins.push({{t:'Top income source',v:topSend.entity,d:'KES '+fmt(topSend.total_amount)+' received across '+topSend.txn_count+' transactions.',c:''}});

  document.getElementById('insightsList').innerHTML=ins.map(i=>`
    <div class="finding ${{i.c}}">
      <div class="finding-title">${{i.t}}</div>
      <div class="finding-val">${{i.v}}</div>
      <div class="finding-desc">${{i.d}}</div>
    </div>`).join('');
}}

// ── ANOMALIES ─────────────────────────────────────────────────────────────────
function renderAnomalies(){{
  const an=RAW.anomalies;
  document.getElementById('anomKpi').innerHTML=[
    kpiCard('Flagged transactions',an.length,null,an.length>5?'amber':''),
    kpiCard('Method','Z-score > 2.5',null,'muted'),
  ].join('');

  document.getElementById('anomTbl').innerHTML=an.map(r=>`<tr>
    <td>${{r.date}}</td><td>${{r.time}}</td>
    <td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{r.details}}</td>
    <td><span class="badge badge-anom">${{r.transaction_type}}</span></td>
    <td style="color:#1baf7a">${{r.paid_in>0?'KES '+fmt(r.paid_in):'—'}}</td>
    <td style="color:#e34948">${{r.withdrawn<0?'KES '+fmt(Math.abs(r.withdrawn)):'—'}}</td>
    <td>KES ${{fmt(r.balance)}}</td>
    <td style="color:#7c6fcd">${{r.z_score.toFixed(2)}}</td>
  </tr>`).join('');
}}

// ── TRANSACTIONS TABLE ─────────────────────────────────────────────────────────
let txnPage_=1; const TXN_PAGE=20; let filteredTxns=[];
function initTxnFilters(){{
  const months=[...new Set(RAW.all_txns.map(r=>r.date.substring(0,7)))].sort();
  const types=[...new Set(RAW.all_txns.map(r=>r.transaction_type))].sort();
  const mSel=document.getElementById('fMonth');
  months.forEach(m=>{{const o=document.createElement('option');o.value=m;o.textContent=m;mSel.appendChild(o);}});
  const tSel=document.getElementById('fType');
  types.forEach(t=>{{const o=document.createElement('option');o.value=t;o.textContent=t;tSel.appendChild(o);}});
}}
function filterTxns(){{
  const mon=document.getElementById('fMonth').value;
  const typ=document.getElementById('fType').value;
  const sta=document.getElementById('fStatus').value;
  const srch=document.getElementById('fSearch').value.toLowerCase();
  filteredTxns=RAW.all_txns.filter(r=>{{
    if(mon&&!r.date.startsWith(mon))return false;
    if(typ&&r.transaction_type!==typ)return false;
    if(sta&&r.status!==sta)return false;
    if(srch&&!r.details.toLowerCase().includes(srch))return false;
    return true;
  }});
  txnPage_=1; renderTxnPage();
}}
function renderTxnPage(){{
  const total=filteredTxns.length;
  const pages=Math.ceil(total/TXN_PAGE)||1;
  if(txnPage_>pages)txnPage_=pages;
  const slice=filteredTxns.slice((txnPage_-1)*TXN_PAGE,txnPage_*TXN_PAGE);
  document.getElementById('txnCountLabel').textContent=total+' records';
  document.getElementById('txnPageInfo').textContent='Page '+txnPage_+' of '+pages;
  document.getElementById('txnTblBody').innerHTML=slice.map((r,i)=>`<tr>
    <td style="color:var(--muted)">${{(txnPage_-1)*TXN_PAGE+i+1}}</td>
    <td style="white-space:nowrap">${{r.date}}</td>
    <td style="white-space:nowrap">${{r.time}}</td>
    <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{r.details}}</td>
    <td><span class="badge badge-${{r.paid_in>0?'in':'out'}}">${{r.transaction_type}}</span></td>
    <td><span class="badge badge-${{r.status==='Completed'?'ok':r.status==='Reversed'?'rev':'fail'}}">${{r.status}}</span></td>
    <td style="text-align:right;color:#1baf7a">${{r.paid_in>0?'KES '+fmt(r.paid_in):'—'}}</td>
    <td style="text-align:right;color:#e34948">${{r.withdrawn<0?'KES '+fmt(Math.abs(r.withdrawn)):'—'}}</td>
    <td style="text-align:right">KES ${{fmt(r.balance)}}</td>
    <td>${{r.is_anomaly?'<span class="badge badge-anom">&#9888; Unusual</span>':''}}</td>
  </tr>`).join('');
}}
function txnPage(d){{txnPage_+=d;if(txnPage_<1)txnPage_=1;renderTxnPage();}}
function resetTxnFilters(){{
  ['fMonth','fType','fStatus'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('fSearch').value='';
  filterTxns();
}}

// ── INIT ──────────────────────────────────────────────────────────────────────
document.getElementById('genAt').textContent='Generated '+RAW.generated_at;
document.getElementById('genInfo').textContent='Dashboard generated '+RAW.generated_at+' — data sourced directly from M-PESA PDF statements.';
renderOverview();
</script>
</body>
</html>"""

    output_path.write_text(html, encoding="utf-8")
    print(f"[SUCCESS] Dashboard saved → {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 — MASTER ENTRY POINT  (called from Part 1 main())
# ─────────────────────────────────────────────────────────────────────────────

def run_analytics(df: pd.DataFrame, pdf_folder: Path) -> None:
    """
    Call this at the end of Part 1's main(), passing the clean df and PDF_FOLDER.

    Example (add to bottom of Part 1 main, after clean_dataframe returns):
        from mpesa_analytics_part2 import run_analytics
        run_analytics(df, PDF_FOLDER)
    """
    print("\n" + "=" * 60)
    print("  M-PESA ANALYTICS ENGINE — PART 2")
    print("=" * 60)

    out_dir = pdf_folder / "mpesa_output"
    out_dir.mkdir(exist_ok=True)

    # 1. Feature engineering
    print("[INFO] Engineering features …")
    df = engineer_features(df)

    # 2. Anomaly detection
    print("[INFO] Running anomaly detection …")
    df = detect_anomalies(df)

    # 3. Entity extraction
    print("[INFO] Extracting sender / recipient entities …")
    top_recipients, top_senders = extract_entities(df)
    print(f"       Top recipients found : {len(top_recipients)}")
    print(f"       Top senders found    : {len(top_senders)}")

    # 4. Core financial metrics
    print("[INFO] Computing financial metrics …")
    metrics = compute_metrics(df)

    # 5. Print summary to console
    m  = metrics
    print("\n── Financial summary ────────────────────────────────────")
    print(f"  Total paid in     : KES {m['income']['total']:>14,.2f}")
    print(f"  Total withdrawn   : KES {m['expense']['total']:>14,.2f}")
    print(f"  Net cash flow     : KES {m['savings']['net_flow']:>14,.2f}")
    print(f"  Savings rate      : {m['savings']['savings_rate']:>10.2f}%")
    print(f"  Health score      : {m['health_score']:>10.1f} / 100")
    print(f"  Best month        : {m['best_month']}")
    print(f"  Worst month       : {m['worst_month']}")
    print(f"  Anomalies flagged : {df['is_anomaly'].sum()}")
    print("─" * 55)

    # 6. Export CSV
    csv_path = out_dir / "mpesa_clean.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"[SUCCESS] CSV saved → {csv_path}")

    # 7. Export Excel
    xlsx_path = out_dir / "mpesa_clean.xlsx"
    export_excel(df, metrics, top_recipients, top_senders, xlsx_path)

    # 8. Build dashboard
    print("[INFO] Building interactive dashboard …")
    dash_path = out_dir / "mpesa_dashboard.html"
    build_dashboard(df, metrics, top_recipients, top_senders, dash_path)

    print("\n── Outputs ──────────────────────────────────────────────")
    print(f"  CSV         : {csv_path}")
    print(f"  Excel       : {xlsx_path}")
    print(f"  Dashboard   : {dash_path}")
    print("─" * 55)
    print("[DONE] Open mpesa_dashboard.html in any browser.\n")
