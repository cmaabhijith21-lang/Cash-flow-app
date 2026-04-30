from __future__ import annotations

# ─────────────────────────────────────────────────────────────────────────────
#  Cash Flow Manager — CFO Edition
#  Improvements over original:
#    P1: Configurable collection probabilities · Min-cash alert banner ·
#        Multi-region holiday calendar (Karnataka 2026-27, Maharashtra,
#        Delhi NCR) · TDS deduction on collections · Typo fixes
#    P2: OD / credit-line KPIs (sanctioned limit, headroom, interest est.) ·
#        Audit trail (load timestamp + file context)
#    P3: CFO narrative commentary boxes · Variance commentary in Actual vs
#        Budget · Commentary included in Excel export
# ─────────────────────────────────────────────────────────────────────────────

import ast
import hashlib
from calendar import monthrange
from datetime import date, datetime, timedelta
from io import BytesIO
from numbers import Number
from pathlib import Path
import re
from typing import Any

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


APP_TITLE = "Cash Flow Manager"
DEFAULT_SOURCE = Path(__file__).with_name("Data Source.xlsx")
TODAY = pd.Timestamp(date.today()).normalize()
DATE_DISPLAY_FORMAT = "%d-%m-%Y"

# ── Default collection probabilities (overridden via sidebar sliders) ─────────
DEFAULT_BASE_COLLECTION_PROBABILITIES: dict[str, float] = {
    "0-30": 1.00,
    "30-60": 1.00,
    "60+": 1.00,
}
DEFAULT_WORST_COLLECTION_PROBABILITIES: dict[str, float] = {
    "0-30": 0.80,
    "30-60": 0.80,
    "60+": 0.80,
}
COLLECTION_DATE_OFFSETS: dict[str, int] = {
    "0-30": 10,
    "30-60": 35,
    "60+": 70,
}

# ── Default model assumption constants ────────────────────────────────────────
DEFAULT_MIN_CASH_THRESHOLD: float = 500_000.0   # ₹5 lakh minimum cash floor
DEFAULT_TDS_RATE: float = 0.0                    # TDS % on gross collections
DEFAULT_OD_LIMIT: float = 0.0                    # Sanctioned OD / CC limit
DEFAULT_OD_RATE_PA: float = 12.0                 # OD interest rate % p.a.

# ── Holiday calendar ──────────────────────────────────────────────────────────
STATE_BANK_HOLIDAYS: dict[str, dict[int, dict[str, str]]] = {
    "Karnataka (Bengaluru)": {
        2026: {
            "2026-01-14": "Makara Sankranti / Uttarayana Punyakala",
            "2026-01-26": "Republic Day",
            "2026-03-20": "Ugadi",
            "2026-03-30": "Mahavir Jayanti (Karnataka observed date)",
            "2026-04-01": "Annual Accounts Closing (Bank Holiday)",
            "2026-04-03": "Good Friday",
            "2026-04-14": "Dr. B. R. Ambedkar Jayanti",
            "2026-04-20": "Basava Jayanthi / Akshaya Tritiya",
            "2026-05-01": "May Day / Buddha Purnima",
            "2026-05-27": "Bakrid / Eid al-Adha - tentative",
            "2026-06-26": "Muharram - tentative",
            "2026-08-15": "Independence Day",
            "2026-08-25": "Id-E-Milad - tentative",
            "2026-09-14": "Ganesh Chaturthi",
            "2026-10-02": "Mahatma Gandhi Jayanti",
            "2026-10-20": "Maha Navami / Ayudha Puja / Vijayadashami",
            "2026-10-26": "Maharshi Valmiki Jayanti",
            "2026-11-09": "Deepavali Holiday",
            "2026-11-27": "Kanakadasa Jayanthi",
            "2026-12-25": "Christmas Day",
        },
        2027: {
            "2027-01-14": "Makara Sankranti",
            "2027-01-26": "Republic Day",
            "2027-04-01": "Annual Accounts Closing (Bank Holiday)",
            "2027-04-14": "Dr. B. R. Ambedkar Jayanti",
            "2027-04-23": "Good Friday - tentative",
            "2027-05-01": "May Day / Basava Jayanthi",
            "2027-08-15": "Independence Day",
            "2027-09-02": "Ganesh Chaturthi - tentative",
            "2027-10-02": "Mahatma Gandhi Jayanti",
            "2027-10-20": "Vijaya Dashami - tentative",
            "2027-11-05": "Deepavali - tentative",
            "2027-12-25": "Christmas Day",
        },
    },
    "Maharashtra (Mumbai)": {
        2026: {
            "2026-01-26": "Republic Day",
            "2026-02-19": "Chhatrapati Shivaji Maharaj Jayanti",
            "2026-03-17": "Holi (Second Day)",
            "2026-03-20": "Gudi Padwa",
            "2026-04-01": "Annual Accounts Closing (Bank Holiday)",
            "2026-04-03": "Good Friday",
            "2026-04-14": "Dr. B. R. Ambedkar Jayanti",
            "2026-05-01": "Maharashtra Day / Buddha Purnima",
            "2026-05-27": "Eid-Ul-Adha - tentative",
            "2026-06-26": "Muharram - tentative",
            "2026-08-15": "Independence Day",
            "2026-08-25": "Id-E-Milad - tentative",
            "2026-09-14": "Ganesh Chaturthi",
            "2026-10-02": "Mahatma Gandhi Jayanti",
            "2026-10-20": "Dussehra",
            "2026-11-09": "Deepavali Holiday",
            "2026-12-25": "Christmas Day",
        },
        2027: {
            "2027-01-26": "Republic Day",
            "2027-02-19": "Chhatrapati Shivaji Maharaj Jayanti",
            "2027-04-01": "Annual Accounts Closing (Bank Holiday)",
            "2027-04-14": "Dr. B. R. Ambedkar Jayanti",
            "2027-05-01": "Maharashtra Day",
            "2027-08-15": "Independence Day",
            "2027-10-02": "Mahatma Gandhi Jayanti",
            "2027-12-25": "Christmas Day",
        },
    },
    "Delhi NCR": {
        2026: {
            "2026-01-26": "Republic Day",
            "2026-02-26": "Maha Shivaratri - tentative",
            "2026-03-17": "Holi",
            "2026-04-01": "Annual Accounts Closing (Bank Holiday)",
            "2026-04-03": "Good Friday",
            "2026-04-14": "Dr. B. R. Ambedkar Jayanti",
            "2026-04-20": "Ram Navami - tentative",
            "2026-05-01": "Buddha Purnima",
            "2026-05-27": "Eid-Ul-Adha - tentative",
            "2026-06-26": "Muharram - tentative",
            "2026-08-15": "Independence Day",
            "2026-08-25": "Id-E-Milad - tentative",
            "2026-10-02": "Mahatma Gandhi Jayanti",
            "2026-10-20": "Dussehra",
            "2026-10-27": "Diwali - tentative",
            "2026-11-05": "Guru Nanak Jayanti - tentative",
            "2026-12-25": "Christmas Day",
        },
        2027: {
            "2027-01-26": "Republic Day",
            "2027-04-01": "Annual Accounts Closing (Bank Holiday)",
            "2027-08-15": "Independence Day",
            "2027-10-02": "Mahatma Gandhi Jayanti",
            "2027-12-25": "Christmas Day",
        },
    },
}

TOTAL_ROWS: set[str] = {
    "Opening Balance",
    "Collections",
    "Total Inflows",
    "Cash Outflows",
    "Total Outflows",
    "Net Movement of Cash",
    "Ending Cash Balance (Planned)",
    "Ending Cash Balance (Actual)",
    "Unexpected Spends / Savings",
}

PREFERRED_OUTFLOW_ORDER: list[str] = [
    "Employee Salaries",
    "Mcom Salaries",
    "Statutory Dues",
    "Fixed Expenses",
    "Vendor Payment",
    "Reimbursements",   # Fixed: was "Reimbursments"
]

CELL_REF_PATTERN = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")

ALLOWED_AST_BINARY_OPERATORS: tuple[type[ast.operator], ...] = (
    ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Mod, ast.Pow,
)
ALLOWED_AST_UNARY_OPERATORS: tuple[type[ast.unaryop], ...] = (
    ast.UAdd, ast.USub,
)


# ══════════════════════════════════════════════════════════════════════════════
#  UTILITY HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def standardize_label(value: str) -> str:
    return "".join(ch.lower() for ch in str(value).strip() if ch.isalnum())


def order_outflow_sheets(sheet_names: list[str]) -> list[str]:
    preferred_rank = {standardize_label(name): idx for idx, name in enumerate(PREFERRED_OUTFLOW_ORDER)}
    return sorted(
        sheet_names,
        key=lambda name: (preferred_rank.get(standardize_label(name), len(preferred_rank)), standardize_label(name)),
    )


def find_column(columns: list[str], options: list[str]) -> str | None:
    normalized = {col: standardize_label(col) for col in columns}
    for option in options:
        target = standardize_label(option)
        for col, label in normalized.items():
            if target and target in label:
                return col
    return None


def to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def to_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.normalize()


def build_summary_label(
    df: pd.DataFrame,
    preferred_column: str | None,
    fallback_column: str | None,
    default_value: str,
) -> pd.Series:
    label = pd.Series("", index=df.index, dtype="object")
    for column in [preferred_column, fallback_column]:
        if column and column in df.columns:
            values = df[column].fillna("").astype(str).str.strip()
            label = label.mask(label.eq(""), values)
    return label.mask(label.eq(""), default_value)


def reset_source_pointer(source: Any) -> None:
    if hasattr(source, "seek"):
        source.seek(0)


def build_month_start(value: Any) -> pd.Timestamp:
    return pd.Timestamp(value).to_period("M").to_timestamp()


def build_month_end(value: Any) -> pd.Timestamp:
    return build_month_start(value) + pd.offsets.MonthEnd(1)


def format_date(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return pd.Timestamp(value).strftime(DATE_DISPLAY_FORMAT)


def month_label(value: Any) -> str:
    return pd.Timestamp(value).strftime("%b %Y")


def default_balance_as_of_date(selected_month: pd.Timestamp) -> pd.Timestamp:
    month_start = build_month_start(selected_month)
    month_end = build_month_end(selected_month)
    if month_start <= TODAY <= month_end:
        return TODAY
    return month_start


def short_date_label(value: pd.Timestamp) -> str:
    return format_date(value)


def full_date_with_day(value: pd.Timestamp) -> str:
    return format_date(value)


def format_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    display_df = df.copy()
    for column in display_df.columns:
        series = display_df[column]
        if pd.api.types.is_datetime64_any_dtype(series):
            display_df[column] = series.map(format_date)
            continue
        if series.dtype == "object":
            non_null = series.dropna()
            if not non_null.empty and non_null.map(
                lambda value: isinstance(value, (pd.Timestamp, datetime, date))
            ).all():
                display_df[column] = series.map(format_date)
    return display_df


def normalize_billing_bucket(value: Any) -> str:
    label = str(value or "").strip().lower()
    if "adv" in label:
        return "Advance"
    return "Arrear"


def file_md5(source: Any) -> str:
    """Return a short MD5 hex digest of the uploaded file bytes."""
    try:
        reset_source_pointer(source)
        digest = hashlib.md5(source.read()).hexdigest()[:8].upper()
        reset_source_pointer(source)
        return digest
    except Exception:
        return "N/A"


# ══════════════════════════════════════════════════════════════════════════════
#  HOLIDAY CALENDAR
# ══════════════════════════════════════════════════════════════════════════════

def get_bank_holiday_events(selected_month: pd.Timestamp, holiday_region: str) -> list[dict[str, Any]]:
    month_start = build_month_start(selected_month)
    month_end = month_start + pd.offsets.MonthEnd(1)
    year = month_start.year
    events: dict[pd.Timestamp, list[str]] = {}
    current = month_start
    saturday_count = 0
    while current <= month_end:
        day_events: list[str] = []
        if current.weekday() == 5:
            saturday_count += 1
            if saturday_count == 2:
                day_events.append("Second Saturday")
            elif saturday_count == 4:
                day_events.append("Fourth Saturday")
        public_holiday_name = (
            STATE_BANK_HOLIDAYS.get(holiday_region, {}).get(year, {}).get(current.strftime("%Y-%m-%d"))
        )
        if public_holiday_name:
            day_events.append(public_holiday_name)
        if day_events:
            events[current] = day_events
        current += pd.Timedelta(days=1)
    return [
        {"date": event_date, "name": " / ".join(names)}
        for event_date, names in sorted(events.items(), key=lambda item: item[0])
    ]


def build_week_ranges(selected_month: pd.Timestamp, holiday_region: str) -> list[dict[str, Any]]:
    start = build_month_start(selected_month)
    end = start + pd.offsets.MonthEnd(1)
    holiday_events = get_bank_holiday_events(selected_month, holiday_region)
    weeks = []
    week_start = start
    idx = 1
    while week_start <= end:
        week_end = min(week_start + pd.Timedelta(days=6), end)
        week_holidays = [e for e in holiday_events if week_start <= e["date"] <= week_end]
        weeks.append({
            "key": f"Week {idx}",
            "label": f"Week {idx}: {full_date_with_day(week_start)} - {full_date_with_day(week_end)}",
            "start": week_start,
            "end": week_end,
            "holidays": week_holidays,
        })
        week_start = week_end + pd.Timedelta(days=1)
        idx += 1
    return weeks


# ══════════════════════════════════════════════════════════════════════════════
#  FORMULA EVALUATION
# ══════════════════════════════════════════════════════════════════════════════

def classify_aging_bucket(ageing: float) -> str:
    if pd.isna(ageing) or ageing <= 30:
        return "0-30"
    if ageing <= 60:
        return "30-60"
    return "60+"


def is_date_like(value: Any) -> bool:
    return isinstance(value, (pd.Timestamp, date)) or hasattr(value, "year")


def expand_cell_range(range_ref: str) -> list[str]:
    start_ref, end_ref = [part.strip().replace("$", "") for part in range_ref.split(":", 1)]
    start_match = CELL_REF_PATTERN.fullmatch(start_ref)
    end_match = CELL_REF_PATTERN.fullmatch(end_ref)
    if not start_match or not end_match:
        return []
    start_col, start_row = start_match.groups()
    end_col, end_row = end_match.groups()
    refs = []
    for row_idx in range(int(start_row), int(end_row) + 1):
        for col_idx in range(
            column_index_from_string(start_col),
            column_index_from_string(end_col) + 1,
        ):
            refs.append(f"{get_column_letter(col_idx)}{row_idx}")
    return refs


def numeric_or_zero(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float, np.number)):
        return float(value)
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def safe_numeric_eval(expression: str) -> float | None:
    try:
        parsed = ast.parse(expression, mode="eval")
    except SyntaxError:
        return None

    def evaluate_node(node: ast.AST) -> float:
        if isinstance(node, ast.Expression):
            return evaluate_node(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return float(node.value)
        if isinstance(node, ast.Num):
            return float(node.n)
        if isinstance(node, ast.BinOp) and isinstance(node.op, ALLOWED_AST_BINARY_OPERATORS):
            left = evaluate_node(node.left)
            right = evaluate_node(node.right)
            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            if isinstance(node.op, ast.Mult):
                return left * right
            if isinstance(node.op, ast.Div):
                return left / right
            if isinstance(node.op, ast.Mod):
                return left % right
            return left ** right
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ALLOWED_AST_UNARY_OPERATORS):
            operand = evaluate_node(node.operand)
            return operand if isinstance(node.op, ast.UAdd) else -operand
        raise ValueError("Unsupported expression")

    try:
        return evaluate_node(parsed)
    except (ValueError, TypeError, ZeroDivisionError, OverflowError):
        return None


def read_workbook_sheets(source: Any) -> dict[str, pd.DataFrame]:
    reset_source_pointer(source)
    workbook_values = load_workbook(source, data_only=True)
    reset_source_pointer(source)
    workbook_formulas = load_workbook(source, data_only=False)
    sheet_frames: dict[str, pd.DataFrame] = {}

    for sheet_name in workbook_values.sheetnames:
        ws_values = workbook_values[sheet_name]
        ws_formulas = workbook_formulas[sheet_name]
        resolved_cache: dict[str, Any] = {}

        def resolve_cell(ref: str, visiting: set[str] | None = None) -> Any:
            cell_ref = ref.replace("$", "")
            if cell_ref in resolved_cache:
                return resolved_cache[cell_ref]
            visiting = visiting or set()
            if cell_ref in visiting:
                return None
            cached_value = ws_values[cell_ref].value
            if cached_value is not None:
                resolved_cache[cell_ref] = cached_value
                return cached_value
            formula_value = ws_formulas[cell_ref].value
            if isinstance(formula_value, str) and formula_value.startswith("="):
                visiting.add(cell_ref)
                resolved = evaluate_formula(formula_value[1:], resolve_cell, visiting)
                visiting.discard(cell_ref)
                resolved_cache[cell_ref] = resolved
                return resolved
            resolved_cache[cell_ref] = formula_value
            return formula_value

        rows: list[list[Any]] = []
        for row_idx in range(1, ws_formulas.max_row + 1):
            row_values = []
            has_content = False
            for col_idx in range(1, ws_formulas.max_column + 1):
                ref = f"{get_column_letter(col_idx)}{row_idx}"
                value = resolve_cell(ref)
                row_values.append(value)
                if value not in (None, "") and not (isinstance(value, float) and pd.isna(value)):
                    has_content = True
            if has_content:
                rows.append(row_values)

        if not rows:
            sheet_frames[sheet_name] = pd.DataFrame()
            continue
        header = []
        for idx, value in enumerate(rows[0], start=1):
            if value in (None, ""):
                header.append(f"Column {idx}")
            else:
                header.append(str(value).strip())
        sheet_frames[sheet_name] = pd.DataFrame(rows[1:], columns=header)

    return sheet_frames


def evaluate_formula(expression: str, resolver: Any, visiting: set[str]) -> Any:
    expr = expression.strip()
    direct_ref = CELL_REF_PATTERN.fullmatch(expr.replace("$", ""))
    if direct_ref:
        return resolver(direct_ref.group(0), visiting)

    date_math = re.fullmatch(r"([A-Z]{1,3}\d+)\s*([+-])\s*(\d+)", expr, flags=re.IGNORECASE)
    if date_math:
        ref, operator, day_count = date_math.groups()
        base_value = resolver(ref, visiting)
        if is_date_like(base_value):
            base_timestamp = pd.Timestamp(base_value)
            delta = pd.Timedelta(days=int(day_count))
            return base_timestamp + delta if operator == "+" else base_timestamp - delta

    def replace_sum(match: re.Match[str]) -> str:
        total = 0.0
        parts = [part.strip() for part in match.group(1).split(",") if part.strip()]
        for part in parts:
            if ":" in part:
                refs = expand_cell_range(part)
                total += sum(numeric_or_zero(resolver(r, visiting)) for r in refs)
            elif CELL_REF_PATTERN.fullmatch(part.replace("$", "")):
                total += numeric_or_zero(resolver(part, visiting))
            else:
                total += numeric_or_zero(evaluate_formula(part, resolver, visiting))
        return str(total)

    expr = re.sub(r"(?i)SUM\(([^)]+)\)", replace_sum, expr)
    expr = re.sub(r"(\d+(?:\.\d+)?)%", lambda m: f"({m.group(1)}/100)", expr)

    def replace_cell(match: re.Match[str]) -> str:
        ref = match.group(0)
        value = resolver(ref, visiting)
        if is_date_like(value):
            return "0"
        return str(numeric_or_zero(value))

    expr = CELL_REF_PATTERN.sub(replace_cell, expr)
    return safe_numeric_eval(expr)


# ══════════════════════════════════════════════════════════════════════════════
#  DATA NORMALISATION
# ══════════════════════════════════════════════════════════════════════════════

def normalize_receivables(df: pd.DataFrame, actual_mode: bool = False) -> pd.DataFrame:
    empty_cols = [
        "invoice_date", "cash_date", "tentative_collection_date", "due_date",
        "billing_type", "billing_bucket", "description", "counterparty",
        "amount", "ageing", "aging_bucket",
    ]
    if df.empty:
        return pd.DataFrame(columns=empty_cols)

    columns = list(df.columns)
    date_col = find_column(columns, ["collection date", "actual date", "date"])
    amount_col = find_column(columns, ["net receipt", "pending amt", "gross amt", "basic amt", "amount"])
    client_col = find_column(columns, ["clients", "customer", "party", "particulars", "particualrs"])
    ref_col = find_column(columns, ["ref no", "invoice no", "invoice number", "inv number"])
    ageing_col = find_column(columns, ["ageing", "aging"])
    tentative_date_col = find_column(columns, ["tentative collection date", "tentative date", "expected collection date"])
    due_date_col = find_column(columns, ["due date"])
    billing_type_col = find_column(columns, ["billing type"])

    if amount_col is None:
        return pd.DataFrame(columns=empty_cols)

    invoice_date = to_datetime(df[date_col]) if date_col else pd.Series(pd.NaT, index=df.index)
    amount = to_numeric(df[amount_col])
    counterparty = df[client_col].fillna("Receivable") if client_col else pd.Series("Receivable", index=df.index)
    reference = df[ref_col].fillna("") if ref_col else pd.Series("", index=df.index)
    description = np.where(
        reference.astype(str).str.strip() != "",
        counterparty.astype(str) + " | " + reference.astype(str),
        counterparty.astype(str),
    )
    ageing = to_numeric(df[ageing_col]) if ageing_col else pd.Series(0, index=df.index, dtype=float)
    aging_bucket = ageing.apply(classify_aging_bucket)
    cash_date = invoice_date if actual_mode else pd.Series(pd.NaT, index=df.index)
    tentative_collection_date = to_datetime(df[tentative_date_col]) if tentative_date_col else pd.Series(pd.NaT, index=df.index)
    due_date = to_datetime(df[due_date_col]) if due_date_col else pd.Series(pd.NaT, index=df.index)
    billing_type = df[billing_type_col].fillna("").astype(str).str.strip() if billing_type_col else pd.Series("", index=df.index, dtype="object")
    billing_bucket = billing_type.map(normalize_billing_bucket)

    normalized = pd.DataFrame({
        "invoice_date": invoice_date,
        "cash_date": cash_date,
        "tentative_collection_date": tentative_collection_date,
        "due_date": due_date,
        "billing_type": billing_type,
        "billing_bucket": billing_bucket,
        "description": description,
        "counterparty": counterparty,
        "amount": amount,
        "ageing": ageing,
        "aging_bucket": aging_bucket,
    })
    normalized = normalized.loc[normalized["amount"] != 0].copy()
    return normalized.reset_index(drop=True)


def normalize_outflow_sheet(df: pd.DataFrame, sheet_name: str, actual_mode: bool = False) -> pd.DataFrame:
    empty_cols = ["sheet_name", "date", "description", "vendor_name", "amount"]
    if df.empty:
        return pd.DataFrame(columns=empty_cols)

    columns = list(df.columns)
    amount_col = find_column(columns, ["payable", "amount", "balance", "gross", "net amount", "net payable"])
    if amount_col is None:
        return pd.DataFrame(columns=empty_cols)

    if actual_mode:
        actual_date_col = find_column(columns, ["actual date"])
        planned_date_col = find_column(columns, ["payment date", "plan payment date", "date", "invoice date"])
        if actual_date_col and planned_date_col and actual_date_col != planned_date_col:
            date_values = to_datetime(df[actual_date_col]).fillna(to_datetime(df[planned_date_col]))
        elif actual_date_col:
            date_values = to_datetime(df[actual_date_col])
        elif planned_date_col:
            date_values = to_datetime(df[planned_date_col])
        else:
            date_values = pd.Series(pd.NaT, index=df.index)
    else:
        date_col = find_column(columns, ["payment date", "plan payment date", "date", "invoice date"])
        date_values = to_datetime(df[date_col]) if date_col else pd.Series(pd.NaT, index=df.index)

    vendor_col = find_column(columns, ["vendors name as per tally", "vendor name", "vendor", "supplier", "payee", "party"])
    description_col = find_column(
        columns,
        ["expense", "nature of payment", "particulars", "particualrs", "invoice no", "inv. number", "inv number"],
    )

    if date_values.isna().all():
        return pd.DataFrame(columns=empty_cols)

    vendor_name = build_summary_label(df, vendor_col, description_col, sheet_name)
    description = build_summary_label(df, description_col, vendor_col, sheet_name)

    normalized = pd.DataFrame({
        "sheet_name": sheet_name,
        "date": date_values,
        "description": description,
        "vendor_name": vendor_name,
        "amount": to_numeric(df[amount_col]),
    })
    normalized = normalized.loc[normalized["date"].notna() & (normalized["amount"] != 0)].copy()
    return normalized.reset_index(drop=True)


def load_data(source: Any, actual_mode: bool = False) -> dict[str, Any]:
    receivables = pd.DataFrame()
    outflow_frames: list[pd.DataFrame] = []
    outflow_sheet_order: list[str] = []
    workbook_sheets = read_workbook_sheets(source)

    for sheet_name, raw in workbook_sheets.items():
        raw = raw.dropna(how="all")
        if raw.empty:
            continue
        if "receivable" in sheet_name.lower():
            receivables = normalize_receivables(raw, actual_mode=actual_mode)
        else:
            normalized = normalize_outflow_sheet(raw, sheet_name, actual_mode=actual_mode)
            if not normalized.empty:
                outflow_frames.append(normalized)
                outflow_sheet_order.append(sheet_name)

    outflows = (
        pd.concat(outflow_frames, ignore_index=True)
        if outflow_frames
        else pd.DataFrame(columns=["sheet_name", "date", "description", "vendor_name", "amount"])
    )
    outflow_sheet_order = order_outflow_sheets(outflow_sheet_order)
    month_candidates = []
    if not outflows.empty:
        month_candidates.extend(outflows["date"].dropna().tolist())
    if actual_mode and not receivables.empty:
        month_candidates.extend(receivables["cash_date"].dropna().tolist())

    available_months = sorted({build_month_start(item) for item in month_candidates})
    return {
        "receivables": receivables,
        "outflows": outflows,
        "outflow_sheet_order": outflow_sheet_order,
        "available_months": available_months,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  COLLECTION PROBABILITY ENGINE  (now configurable)
# ══════════════════════════════════════════════════════════════════════════════

def derive_collection_probabilities(
    receivables_df: pd.DataFrame,
    anchor_date: pd.Timestamp | None = None,
    *,
    base_probs: dict[str, float] | None = None,
    worst_probs: dict[str, float] | None = None,
    tds_rate: float = 0.0,
) -> pd.DataFrame:
    """
    Attach collection probability columns to the receivables dataframe.

    Parameters
    ----------
    base_probs / worst_probs : override the module-level defaults per bucket.
    tds_rate : fraction deducted at source (e.g. 0.02 for 2% TDS).
               Applied AFTER probability so expected cash = gross × prob × (1 − tds).
    """
    if receivables_df.empty:
        return receivables_df.copy()

    eff_base = base_probs or DEFAULT_BASE_COLLECTION_PROBABILITIES
    eff_worst = worst_probs or DEFAULT_WORST_COLLECTION_PROBABILITIES
    net_factor = 1.0 - max(0.0, min(tds_rate, 1.0))

    anchor = build_month_start(anchor_date or TODAY)
    derived = receivables_df.copy()
    derived["base_probability"] = derived["aging_bucket"].map(eff_base).fillna(0.0)
    derived["best_probability"] = derived["base_probability"]
    derived["worst_probability"] = derived["aging_bucket"].map(eff_worst).fillna(0.0)
    derived["expected_base"] = derived["amount"] * derived["base_probability"] * net_factor
    derived["expected_best"] = derived["amount"] * derived["best_probability"] * net_factor
    derived["expected_worst"] = derived["amount"] * derived["worst_probability"] * net_factor

    fallback_dates = derived["aging_bucket"].map(
        lambda bucket: anchor + pd.Timedelta(days=COLLECTION_DATE_OFFSETS[bucket])
    )
    if "tentative_collection_date" in derived.columns:
        derived["forecast_collection_date"] = derived["tentative_collection_date"].fillna(fallback_dates)
    else:
        derived["forecast_collection_date"] = fallback_dates
    return derived


def actual_receivables_look_plan_aligned(receivables_df: pd.DataFrame) -> bool:
    required_columns = {"invoice_date", "cash_date", "tentative_collection_date", "ageing"}
    if receivables_df.empty or not required_columns.issubset(receivables_df.columns):
        return False
    if not receivables_df["tentative_collection_date"].notna().any():
        return False
    comparable = receivables_df.loc[
        receivables_df["cash_date"].notna() & receivables_df["invoice_date"].notna(),
        ["cash_date", "invoice_date"],
    ]
    if comparable.empty:
        return False
    return (comparable["cash_date"] == comparable["invoice_date"]).all()


def build_actual_receivables_view(
    actual_data: dict[str, Any],
    anchor_date: pd.Timestamp,
    *,
    base_probs: dict[str, float] | None = None,
    worst_probs: dict[str, float] | None = None,
    tds_rate: float = 0.0,
) -> pd.DataFrame:
    receivables = actual_data.get("receivables", pd.DataFrame()).copy()
    if receivables.empty:
        return receivables
    if actual_receivables_look_plan_aligned(receivables):
        aligned = derive_collection_probabilities(
            receivables,
            anchor_date=anchor_date,
            base_probs=base_probs,
            worst_probs=worst_probs,
            tds_rate=tds_rate,
        )
        aligned["comparison_date"] = aligned["forecast_collection_date"]
        aligned["comparison_amount"] = aligned["expected_base"]
        return aligned
    receivables["comparison_date"] = receivables["cash_date"]
    receivables["comparison_amount"] = receivables["amount"]
    return receivables


# ══════════════════════════════════════════════════════════════════════════════
#  FILTERING & HORIZON HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def filter_between(df: pd.DataFrame, date_column: str, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    return df.loc[df[date_column].between(start, end, inclusive="both")].copy()


def build_month_horizon(selected_month: pd.Timestamp, periods: int = 3) -> list[pd.Timestamp]:
    start = build_month_start(selected_month)
    return [start + pd.DateOffset(months=offset) for offset in range(periods)]


# ══════════════════════════════════════════════════════════════════════════════
#  MONTHLY ROLL-FORWARD
# ══════════════════════════════════════════════════════════════════════════════

def compute_monthly_rollforward(
    receivables_df: pd.DataFrame,
    outflows_df: pd.DataFrame,
    months: list[pd.Timestamp],
    opening_balance: float,
    scenario_column: str = "expected_base",
    period_start_overrides: dict[pd.Timestamp, pd.Timestamp] | None = None,
) -> pd.DataFrame:
    rows = []
    running_open = opening_balance
    for month_start in months:
        effective_start = (period_start_overrides or {}).get(month_start, month_start)
        month_end = month_start + pd.offsets.MonthEnd(1)
        inflows = (
            filter_between(receivables_df, "forecast_collection_date", effective_start, month_end)[scenario_column].sum()
            if not receivables_df.empty
            else 0.0
        )
        outflows = (
            filter_between(outflows_df, "date", effective_start, month_end)["amount"].sum()
            if not outflows_df.empty
            else 0.0
        )
        closing = running_open + inflows - outflows
        rows.append({
            "Month": month_label(month_start),
            "Opening Cash": running_open,
            "Inflows": inflows,
            "Outflows": outflows,
            "Closing Cash": closing,
            "month_start": month_start,
        })
        running_open = closing
    return pd.DataFrame(rows)


def build_monthly_detail_matrix(
    df: pd.DataFrame,
    date_column: str,
    value_column: str,
    line_item_column: str,
    months: list[pd.Timestamp],
    line_item_label: str,
    period_start_overrides: dict[pd.Timestamp, pd.Timestamp] | None = None,
) -> pd.DataFrame:
    month_labels = [month_label(month) for month in months]
    if df.empty:
        return pd.DataFrame(columns=["Line Item"] + month_labels)

    detail = df.copy()
    detail["forecast_month"] = pd.to_datetime(detail[date_column], errors="coerce").dt.to_period("M").dt.to_timestamp()
    detail = detail.loc[detail["forecast_month"].isin(months)].copy()
    detail["period_start"] = detail["forecast_month"].map(
        lambda month: (period_start_overrides or {}).get(month, month)
    )
    detail = detail.loc[pd.to_datetime(detail[date_column], errors="coerce") >= detail["period_start"]].copy()
    if detail.empty:
        return pd.DataFrame(columns=["Line Item"] + month_labels)

    pivot = (
        detail.pivot_table(
            index=line_item_column,
            columns="forecast_month",
            values=value_column,
            aggfunc="sum",
            fill_value=0.0,
        )
        .reindex(columns=months, fill_value=0.0)
        .sort_index()
    )
    pivot.columns = month_labels
    pivot = pivot.reset_index().rename(columns={line_item_column: "Line Item"})
    total_row = {"Line Item": line_item_label}
    for lbl in month_labels:
        total_row[lbl] = pivot[lbl].sum()
    return pd.concat([pd.DataFrame([total_row]), pivot], ignore_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  WEEK DETAIL COMPUTATION
# ══════════════════════════════════════════════════════════════════════════════

def build_week_detail(
    start: pd.Timestamp,
    end: pd.Timestamp,
    plan_receivables: pd.DataFrame,
    plan_outflows: pd.DataFrame,
    outflow_sheet_order: list[str],
    actual_data: dict[str, Any] | None = None,
    effective_start: pd.Timestamp | None = None,
    actual_anchor_date: pd.Timestamp | None = None,
    *,
    base_probs: dict[str, float] | None = None,
    worst_probs: dict[str, float] | None = None,
    tds_rate: float = 0.0,
) -> dict[str, Any]:
    active_start = max(start, effective_start) if effective_start is not None else start
    if active_start > end:
        week_plan_receipts = pd.DataFrame(columns=plan_receivables.columns)
        week_plan_outflows = pd.DataFrame(columns=plan_outflows.columns)
    else:
        week_plan_receipts = filter_between(plan_receivables, "forecast_collection_date", active_start, end)
        week_plan_outflows = filter_between(plan_outflows, "date", active_start, end)

    plan_billing_bucket = (
        week_plan_receipts["billing_bucket"]
        if "billing_bucket" in week_plan_receipts.columns
        else pd.Series("Arrear", index=week_plan_receipts.index)
    )
    plan_advance_collections = (
        week_plan_receipts.loc[plan_billing_bucket == "Advance", "expected_base"].sum()
        if not week_plan_receipts.empty else 0.0
    )
    plan_arrear_collections = (
        week_plan_receipts.loc[plan_billing_bucket != "Advance", "expected_base"].sum()
        if not week_plan_receipts.empty else 0.0
    )
    plan_collections = week_plan_receipts["expected_base"].sum() if not week_plan_receipts.empty else 0.0
    plan_by_category = (
        week_plan_outflows.groupby("sheet_name", as_index=True)["amount"]
        .sum()
        .reindex(outflow_sheet_order, fill_value=0.0)
        if not week_plan_outflows.empty
        else pd.Series(0.0, index=outflow_sheet_order)
    )
    total_outflows = plan_by_category.sum()
    net_plan = plan_collections - total_outflows

    actual_collections = 0.0
    actual_by_category = pd.Series(0.0, index=outflow_sheet_order)
    actual_net = None
    actual_week_receipts = pd.DataFrame()
    actual_week_outflows = pd.DataFrame()
    actual_advance_collections = 0.0
    actual_arrear_collections = 0.0

    if actual_data:
        actual_receivables = build_actual_receivables_view(
            actual_data,
            anchor_date=actual_anchor_date or active_start,
            base_probs=base_probs,
            worst_probs=worst_probs,
            tds_rate=tds_rate,
        )
        actual_outflows_df = actual_data["outflows"]
        if active_start <= end:
            actual_week_receipts = filter_between(actual_receivables, "comparison_date", active_start, end)
            actual_week_outflows = filter_between(actual_outflows_df, "date", active_start, end)
        actual_collections = actual_week_receipts["comparison_amount"].sum() if not actual_week_receipts.empty else 0.0
        actual_billing_bucket = (
            actual_week_receipts["billing_bucket"]
            if "billing_bucket" in actual_week_receipts.columns
            else pd.Series("Arrear", index=actual_week_receipts.index)
        )
        actual_advance_collections = (
            actual_week_receipts.loc[actual_billing_bucket == "Advance", "comparison_amount"].sum()
            if not actual_week_receipts.empty else 0.0
        )
        actual_arrear_collections = (
            actual_week_receipts.loc[actual_billing_bucket != "Advance", "comparison_amount"].sum()
            if not actual_week_receipts.empty else 0.0
        )
        actual_by_category = (
            actual_week_outflows.groupby("sheet_name", as_index=True)["amount"]
            .sum()
            .reindex(outflow_sheet_order, fill_value=0.0)
            if not actual_week_outflows.empty
            else pd.Series(0.0, index=outflow_sheet_order)
        )
        actual_net = actual_collections - actual_by_category.sum()

    return {
        "start": start,
        "end": end,
        "effective_start": active_start,
        "plan_receipts_detail": week_plan_receipts,
        "plan_outflows_detail": week_plan_outflows,
        "plan_advance_collections": plan_advance_collections,
        "plan_arrear_collections": plan_arrear_collections,
        "plan_collections": plan_collections,
        "plan_by_category": plan_by_category,
        "plan_total_outflows": total_outflows,
        "plan_net": net_plan,
        "actual_receipts_detail": actual_week_receipts,
        "actual_outflows_detail": actual_week_outflows,
        "actual_advance_collections": actual_advance_collections,
        "actual_arrear_collections": actual_arrear_collections,
        "actual_collections": actual_collections,
        "actual_by_category": actual_by_category,
        "actual_net": actual_net,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  WEEKLY CASH FLOW COMPUTATION
# ══════════════════════════════════════════════════════════════════════════════

def compute_weekly_cashflow(
    data: dict[str, Any],
    selected_month: pd.Timestamp,
    opening_balance: float,
    holiday_region: str,
    balance_as_of_date: pd.Timestamp | None = None,
    actual_data: dict[str, Any] | None = None,
    *,
    base_probs: dict[str, float] | None = None,
    worst_probs: dict[str, float] | None = None,
    tds_rate: float = 0.0,
    od_sanctioned_limit: float = 0.0,
    od_interest_rate_pa: float = 0.0,
) -> dict[str, Any]:
    month_start = build_month_start(selected_month)
    month_end = build_month_end(selected_month)
    balance_date = pd.Timestamp(balance_as_of_date or month_start).normalize()
    balance_date = min(max(balance_date, month_start), month_end)
    period_start_overrides = {month_start: balance_date}

    plan_receivables = derive_collection_probabilities(
        data["receivables"],
        anchor_date=balance_date,
        base_probs=base_probs,
        worst_probs=worst_probs,
        tds_rate=tds_rate,
    )
    plan_outflows = data["outflows"].copy()
    outflow_sheet_order = data["outflow_sheet_order"]
    weeks = build_week_ranges(month_start, holiday_region)

    horizon_months = build_month_horizon(month_start, periods=3)
    monthly_forecast = compute_monthly_rollforward(
        plan_receivables,
        plan_outflows,
        horizon_months,
        opening_balance,
        scenario_column="expected_base",
        period_start_overrides=period_start_overrides,
    )
    monthly_inflow_detail = build_monthly_detail_matrix(
        plan_receivables,
        date_column="forecast_collection_date",
        value_column="expected_base",
        line_item_column="counterparty",
        months=horizon_months,
        line_item_label="Collections",
        period_start_overrides=period_start_overrides,
    )
    monthly_outflow_detail = build_monthly_detail_matrix(
        plan_outflows,
        date_column="date",
        value_column="amount",
        line_item_column="sheet_name",
        months=horizon_months,
        line_item_label="Cash Outflows",
        period_start_overrides=period_start_overrides,
    )

    running_plan = opening_balance
    running_actual = opening_balance if actual_data else None
    week_meta: list[dict[str, Any]] = []

    for week in weeks:
        week_detail = build_week_detail(
            week["start"],
            week["end"],
            plan_receivables,
            plan_outflows,
            outflow_sheet_order,
            actual_data=actual_data,
            effective_start=balance_date,
            actual_anchor_date=balance_date,
            base_probs=base_probs,
            worst_probs=worst_probs,
            tds_rate=tds_rate,
        )

        # ── OD utilisation & interest for this week ───────────────────────
        days_in_week = max((week["end"] - week["start"]).days + 1, 1)

        if week_detail["effective_start"] > week_detail["end"]:
            week_detail["opening_plan"] = None
            week_detail["ending_plan"] = None
            week_detail["plan_od_utilization"] = None
            week_detail["plan_total_inflows"] = None
            week_detail["plan_od_interest"] = None
        else:
            week_detail["opening_plan"] = running_plan
            raw_od = max(
                week_detail["plan_total_outflows"] - (running_plan + week_detail["plan_collections"]), 0.0
            )
            # Cap OD at sanctioned limit
            capped_od = min(raw_od, od_sanctioned_limit) if od_sanctioned_limit > 0 else raw_od
            week_detail["plan_od_utilization"] = capped_od
            week_detail["plan_total_inflows"] = (
                week_detail["plan_collections"] + week_detail["plan_od_utilization"]
            )
            week_detail["plan_net"] = (
                week_detail["plan_total_inflows"] - week_detail["plan_total_outflows"]
            )
            week_detail["ending_plan"] = running_plan + week_detail["plan_net"]
            # Estimated OD interest for this week
            week_detail["plan_od_interest"] = (
                capped_od * (od_interest_rate_pa / 100.0) / 365.0 * days_in_week
                if od_interest_rate_pa > 0 and capped_od > 0
                else 0.0
            )
            running_plan = week_detail["ending_plan"]

        if actual_data:
            if week_detail["effective_start"] > week_detail["end"]:
                week_detail["opening_actual"] = None
                week_detail["ending_actual"] = None
                week_detail["actual_od_utilization"] = None
                week_detail["actual_total_inflows"] = None
                week_detail["unexpected"] = None
                week_detail["actual_od_interest"] = None
            else:
                week_detail["opening_actual"] = running_actual
                actual_outflow_total = week_detail["actual_by_category"].sum()
                raw_od_actual = max(
                    actual_outflow_total - (running_actual + week_detail["actual_collections"]), 0.0
                )
                capped_od_actual = min(raw_od_actual, od_sanctioned_limit) if od_sanctioned_limit > 0 else raw_od_actual
                week_detail["actual_od_utilization"] = capped_od_actual
                week_detail["actual_total_inflows"] = (
                    week_detail["actual_collections"] + week_detail["actual_od_utilization"]
                )
                actual_net = week_detail["actual_total_inflows"] - actual_outflow_total
                week_detail["actual_net"] = actual_net
                week_detail["ending_actual"] = running_actual + actual_net
                running_actual = week_detail["ending_actual"]
                week_detail["unexpected"] = actual_net - (week_detail["plan_net"] or 0.0)
                week_detail["actual_od_interest"] = (
                    capped_od_actual * (od_interest_rate_pa / 100.0) / 365.0 * days_in_week
                    if od_interest_rate_pa > 0 and capped_od_actual > 0
                    else 0.0
                )
        else:
            week_detail["opening_actual"] = None
            week_detail["ending_actual"] = None
            week_detail["actual_od_utilization"] = None
            week_detail["actual_total_inflows"] = None
            week_detail["unexpected"] = (
                None if week_detail["effective_start"] > week_detail["end"] else 0.0
            )
            week_detail["actual_od_interest"] = None

        week_meta.append(week_detail | {
            "key": week["key"],
            "label": week["label"],
            "holidays": week.get("holidays", []),
        })

    # ── Build table rows ──────────────────────────────────────────────────────
    line_items = [
        "Opening Balance", "Collections", "   Advance", "   Arrear",
        "   OD Utilization", "Total Inflows", "Cash Outflows",
    ]
    line_items.extend([f"   {sheet}" for sheet in outflow_sheet_order])
    line_items.extend([
        "Total Outflows",
        "Net Movement of Cash",
        "Ending Cash Balance (Planned)",
        "Ending Cash Balance (Actual)",
        "Unexpected Spends / Savings",
    ])

    week_values: dict[str, dict[str, Any]] = {item: {} for item in line_items}

    for wd in week_meta:
        key = wd["key"]
        if wd["effective_start"] > wd["end"]:
            for item in line_items:
                week_values[item][key] = None
            continue
        week_values["Opening Balance"][key] = wd["opening_plan"]
        week_values["Collections"][key] = wd["plan_collections"]
        week_values["   Advance"][key] = wd["plan_advance_collections"]
        week_values["   Arrear"][key] = wd["plan_arrear_collections"]
        week_values["   OD Utilization"][key] = wd["plan_od_utilization"]
        week_values["Total Inflows"][key] = wd["plan_total_inflows"]
        week_values["Cash Outflows"][key] = None
        for sheet in outflow_sheet_order:
            week_values[f"   {sheet}"][key] = wd["plan_by_category"].get(sheet, 0.0)
        week_values["Total Outflows"][key] = wd["plan_total_outflows"]
        week_values["Net Movement of Cash"][key] = wd["plan_net"]
        week_values["Ending Cash Balance (Planned)"][key] = wd["ending_plan"]
        week_values["Ending Cash Balance (Actual)"][key] = wd["ending_actual"]
        week_values["Unexpected Spends / Savings"][key] = wd["unexpected"] if actual_data else None

    active_weeks = [wd for wd in week_meta if wd["effective_start"] <= wd["end"]]
    monthly_plan_advance = sum(wd["plan_advance_collections"] for wd in active_weeks)
    monthly_plan_arrear = sum(wd["plan_arrear_collections"] for wd in active_weeks)
    monthly_plan_collections = sum(wd["plan_collections"] for wd in active_weeks)
    monthly_plan_od_utilization = sum(wd["plan_od_utilization"] or 0.0 for wd in active_weeks)
    monthly_plan_total_inflows = sum(wd["plan_total_inflows"] or 0.0 for wd in active_weeks)
    monthly_plan_by_category = (
        pd.DataFrame([wd["plan_by_category"] for wd in active_weeks])
        .fillna(0.0)
        .sum()
        .reindex(outflow_sheet_order, fill_value=0.0)
        if active_weeks
        else pd.Series(0.0, index=outflow_sheet_order)
    )
    monthly_plan_outflows = sum(wd["plan_total_outflows"] for wd in active_weeks)
    monthly_plan_net = sum(wd["plan_net"] for wd in active_weeks)
    monthly_plan_ending = active_weeks[-1]["ending_plan"] if active_weeks else opening_balance
    monthly_plan_od_interest = sum(wd.get("plan_od_interest") or 0.0 for wd in active_weeks)

    # ── Actuals monthly aggregation ───────────────────────────────────────────
    actual_month_collections = actual_month_advance = actual_month_arrear = None
    actual_month_od_utilization = actual_month_total_inflows = None
    actual_month_by_category = pd.Series(dtype=float)
    actual_month_outflows = actual_month_net = actual_month_ending = None
    unexpected_month = None

    if actual_data:
        actual_month_advance = sum((wd.get("actual_advance_collections") or 0.0) for wd in active_weeks)
        actual_month_arrear = sum((wd.get("actual_arrear_collections") or 0.0) for wd in active_weeks)
        actual_month_collections = sum((wd.get("actual_collections") or 0.0) for wd in active_weeks)
        actual_month_od_utilization = sum((wd.get("actual_od_utilization") or 0.0) for wd in active_weeks)
        actual_month_total_inflows = sum((wd.get("actual_total_inflows") or 0.0) for wd in active_weeks)
        actual_month_by_category = (
            pd.DataFrame([wd["actual_by_category"] for wd in active_weeks])
            .fillna(0.0)
            .sum()
            .reindex(outflow_sheet_order, fill_value=0.0)
            if active_weeks
            else pd.Series(0.0, index=outflow_sheet_order)
        )
        actual_month_outflows = actual_month_by_category.sum()
        actual_month_net = sum((wd.get("actual_net") or 0.0) for wd in active_weeks)
        actual_month_ending = active_weeks[-1]["ending_actual"] if active_weeks else opening_balance
        unexpected_month = actual_month_net - monthly_plan_net

    table_rows = []
    for line_item in line_items:
        row = {"Line Item": line_item}
        for week in weeks:
            row[week["key"]] = week_values[line_item].get(week["key"])
        if line_item == "Opening Balance":
            row["Fcst (Selected Month)"] = opening_balance
            row["Actual (if available)"] = opening_balance if actual_data else None
        elif line_item == "Collections":
            row["Fcst (Selected Month)"] = monthly_plan_collections
            row["Actual (if available)"] = actual_month_collections
        elif line_item == "   Advance":
            row["Fcst (Selected Month)"] = monthly_plan_advance
            row["Actual (if available)"] = actual_month_advance
        elif line_item == "   Arrear":
            row["Fcst (Selected Month)"] = monthly_plan_arrear
            row["Actual (if available)"] = actual_month_arrear
        elif line_item == "   OD Utilization":
            row["Fcst (Selected Month)"] = monthly_plan_od_utilization
            row["Actual (if available)"] = actual_month_od_utilization
        elif line_item == "Total Inflows":
            row["Fcst (Selected Month)"] = monthly_plan_total_inflows
            row["Actual (if available)"] = actual_month_total_inflows
        elif line_item == "Cash Outflows":
            row["Fcst (Selected Month)"] = None
            row["Actual (if available)"] = None
        elif line_item.startswith("   "):
            category = line_item[3:]
            row["Fcst (Selected Month)"] = monthly_plan_by_category.get(category, 0.0)
            row["Actual (if available)"] = (
                actual_month_by_category.get(category, 0.0) if actual_data else None
            )
        elif line_item == "Total Outflows":
            row["Fcst (Selected Month)"] = monthly_plan_outflows
            row["Actual (if available)"] = actual_month_outflows
        elif line_item == "Net Movement of Cash":
            row["Fcst (Selected Month)"] = monthly_plan_net
            row["Actual (if available)"] = actual_month_net
        elif line_item == "Ending Cash Balance (Planned)":
            row["Fcst (Selected Month)"] = monthly_plan_ending
            row["Actual (if available)"] = None
        elif line_item == "Ending Cash Balance (Actual)":
            row["Fcst (Selected Month)"] = None
            row["Actual (if available)"] = actual_month_ending
        elif line_item == "Unexpected Spends / Savings":
            row["Fcst (Selected Month)"] = 0.0
            row["Actual (if available)"] = unexpected_month
        table_rows.append(row)

    matrix = pd.DataFrame(table_rows)
    return {
        "weekly_table": matrix,
        "week_meta": week_meta,
        "week_columns": [week["key"] for week in weeks],
        "monthly_forecast": monthly_forecast,
        "monthly_inflow_detail": monthly_inflow_detail,
        "monthly_outflow_detail": monthly_outflow_detail,
        "plan_receivables": plan_receivables,
        "plan_outflows": plan_outflows,
        "selected_month_start": month_start,
        "selected_month_end": month_end,
        "selected_period_start": balance_date,
        "selected_month_plan": {
            "opening": opening_balance,
            "collections": monthly_plan_collections,
            "outflows": monthly_plan_outflows,
            "net": monthly_plan_net,
            "ending": monthly_plan_ending,
            "od_utilization": monthly_plan_od_utilization,
            "od_interest_est": monthly_plan_od_interest,
        },
    }


# ══════════════════════════════════════════════════════════════════════════════
#  COLLECTION ENGINE & SCENARIOS
# ══════════════════════════════════════════════════════════════════════════════

def generate_scenarios(
    receivables_df: pd.DataFrame,
    opening_balance: float,
    next_30_day_outflows: float,
    anchor_date: pd.Timestamp | None = None,
    *,
    base_probs: dict[str, float] | None = None,
    worst_probs: dict[str, float] | None = None,
    tds_rate: float = 0.0,
) -> pd.DataFrame:
    if receivables_df.empty:
        return pd.DataFrame(columns=["Scenario", "Expected Collection", "Impact on Closing Cash", "Runway"])

    derived = derive_collection_probabilities(
        receivables_df, anchor_date=anchor_date,
        base_probs=base_probs, worst_probs=worst_probs, tds_rate=tds_rate,
    )
    horizon_start = build_month_start(anchor_date or TODAY)
    horizon_end = horizon_start + pd.Timedelta(days=29)
    horizon_receivables = filter_between(derived, "forecast_collection_date", horizon_start, horizon_end)
    avg_daily_outflow = next_30_day_outflows / 30 if next_30_day_outflows else 0.0

    rows = []
    for scenario, column in [("Best Case", "expected_best"), ("Base Case", "expected_base"), ("Worst Case", "expected_worst")]:
        expected = horizon_receivables[column].sum() if not horizon_receivables.empty else 0.0
        impact = opening_balance + expected - next_30_day_outflows
        runway = np.inf if avg_daily_outflow == 0 else max(impact, 0.0) / avg_daily_outflow
        rows.append({"Scenario": scenario, "Expected Collection": expected, "Impact on Closing Cash": impact, "Runway": runway})
    return pd.DataFrame(rows)


def calculate_collection_engine(
    receivables: pd.DataFrame,
    selected_month: pd.Timestamp,
    opening_balance: float,
    next_30_day_outflows: float,
    *,
    base_probs: dict[str, float] | None = None,
    worst_probs: dict[str, float] | None = None,
    tds_rate: float = 0.0,
) -> tuple[pd.DataFrame, float, float]:
    derived = derive_collection_probabilities(
        receivables, anchor_date=selected_month,
        base_probs=base_probs, worst_probs=worst_probs, tds_rate=tds_rate,
    )
    scenario_table = generate_scenarios(
        receivables, opening_balance, next_30_day_outflows, anchor_date=selected_month,
        base_probs=base_probs, worst_probs=worst_probs, tds_rate=tds_rate,
    )
    overdue_expected = derived.loc[derived["aging_bucket"].isin(["30-60", "60+"]), "expected_base"].sum() if not derived.empty else 0.0
    sixty_plus_expected = derived.loc[derived["aging_bucket"] == "60+", "expected_base"].sum() if not derived.empty else 0.0
    total_expected = derived["expected_base"].sum() if not derived.empty else 0.0
    overdue_dependency = overdue_expected / total_expected if total_expected else 0.0
    sixty_plus_dependency = sixty_plus_expected / total_expected if total_expected else 0.0
    return scenario_table, overdue_dependency, sixty_plus_dependency


# ══════════════════════════════════════════════════════════════════════════════
#  ACTUAL vs BUDGET COMPARISON
# ══════════════════════════════════════════════════════════════════════════════

def compare_actual_vs_budget(
    weekly_cashflow: dict[str, Any],
    actual_data: dict[str, Any] | None,
) -> pd.DataFrame:
    if not actual_data:
        return pd.DataFrame(columns=["Line Item", "Planned", "Actual", "Variance", "Variance %"])

    table = weekly_cashflow["weekly_table"].copy()
    rows = []
    ending_planned = np.nan
    ending_actual = np.nan
    for _, record in table.iterrows():
        line_item = str(record["Line Item"]).strip()
        if line_item == "Cash Outflows":
            continue
        if line_item == "Ending Cash Balance (Planned)":
            ending_planned = record["Fcst (Selected Month)"]
            continue
        if line_item == "Ending Cash Balance (Actual)":
            ending_actual = record["Actual (if available)"]
            continue
        planned = record["Fcst (Selected Month)"]
        actual = record["Actual (if available)"]
        if pd.isna(planned) and pd.isna(actual):
            continue
        planned_value = 0.0 if pd.isna(planned) else float(planned)
        actual_value = 0.0 if pd.isna(actual) else float(actual)
        variance = actual_value - planned_value
        denominator = abs(planned) if planned not in (None, 0) and not pd.isna(planned) else np.nan
        variance_pct = abs(variance) / denominator if denominator and not pd.isna(denominator) else np.nan
        rows.append({"Line Item": line_item, "Planned": planned, "Actual": actual, "Variance": variance, "Variance %": variance_pct})

    if not pd.isna(ending_planned) or not pd.isna(ending_actual):
        planned_value = 0.0 if pd.isna(ending_planned) else float(ending_planned)
        actual_value = 0.0 if pd.isna(ending_actual) else float(ending_actual)
        variance = actual_value - planned_value
        denominator = abs(ending_planned) if ending_planned not in (None, 0) and not pd.isna(ending_planned) else np.nan
        variance_pct = abs(variance) / denominator if denominator and not pd.isna(denominator) else np.nan
        rows.append({"Line Item": "Ending Cash Balance", "Planned": ending_planned, "Actual": ending_actual, "Variance": variance, "Variance %": variance_pct})
    return pd.DataFrame(rows)


def summarize_line_dates(series: pd.Series) -> str:
    dates = sorted({pd.Timestamp(v).normalize() for v in series.dropna()})
    if not dates:
        return ""
    if len(dates) == 1:
        return format_date(dates[0])
    return f"{format_date(dates[0])} to {format_date(dates[-1])}"


def summarize_inflow_detail(df: pd.DataFrame, *, value_column: str, value_label: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Billing Type", "Client", "Due Window", "Forecast Window", value_label, "Source Rows"])
    working_df = df.copy()
    working_df["_summary_line"] = build_summary_label(working_df, "counterparty", "description", "Collections")
    working_df["_billing_bucket"] = working_df.get("billing_bucket", pd.Series("Arrear", index=working_df.index)).fillna("Arrear")
    return (
        working_df.groupby(["_billing_bucket", "_summary_line"], dropna=False)
        .agg(**{
            value_label: (value_column, "sum"),
            "Due Window": ("due_date", summarize_line_dates),
            "Forecast Window": ("forecast_collection_date", summarize_line_dates),
            "Source Rows": (value_column, "size"),
        })
        .reset_index()
        .rename(columns={"_billing_bucket": "Billing Type", "_summary_line": "Client"})
        .sort_values(by=["Billing Type", value_label, "Client"], ascending=[True, False, True], kind="stable")
        .reset_index(drop=True)
    )


def summarize_outflow_detail(df: pd.DataFrame, *, value_label: str = "Amount") -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Category", "Vendor", "Payment Window", value_label, "Source Rows"])
    working_df = df.copy()
    working_df["_summary_line"] = build_summary_label(working_df, "vendor_name", "description", "Outflows")
    return (
        working_df.groupby(["sheet_name", "_summary_line"], dropna=False)
        .agg(**{
            value_label: ("amount", "sum"),
            "Payment Window": ("date", summarize_line_dates),
            "Source Rows": ("amount", "size"),
        })
        .reset_index()
        .rename(columns={"sheet_name": "Category", "_summary_line": "Vendor"})
        .sort_values(by=["Category", value_label, "Vendor"], ascending=[True, False, True], kind="stable")
        .reset_index(drop=True)
    )


def build_line_level_variance(
    weekly_cashflow: dict[str, Any],
    actual_data: dict[str, Any] | None,
    *,
    base_probs: dict[str, float] | None = None,
    worst_probs: dict[str, float] | None = None,
    tds_rate: float = 0.0,
) -> pd.DataFrame:
    columns = [
        "Section", "Category", "Line Item",
        "Planned Date", "Actual Date", "Status",
        "Planned", "Actual", "Variance", "Variance %",
    ]
    if not actual_data:
        return pd.DataFrame(columns=columns)

    month_start = weekly_cashflow.get("selected_period_start", weekly_cashflow["selected_month_start"])
    month_end = weekly_cashflow["selected_month_end"]

    plan_receivables = filter_between(weekly_cashflow["plan_receivables"], "forecast_collection_date", month_start, month_end)
    actual_receivables_view = build_actual_receivables_view(
        actual_data, anchor_date=month_start,
        base_probs=base_probs, worst_probs=worst_probs, tds_rate=tds_rate,
    )
    actual_receivables = filter_between(actual_receivables_view, "comparison_date", month_start, month_end)
    plan_outflows = filter_between(weekly_cashflow["plan_outflows"], "date", month_start, month_end)
    actual_outflows = filter_between(actual_data["outflows"], "date", month_start, month_end)

    def aggregate_lines(
        df, *, section, category_column, summary_column, summary_fallback_column,
        date_column, value_column, value_label, date_label, category_label=None,
    ):
        if df.empty:
            return pd.DataFrame(columns=["Section", "Category", "Line Item", date_label, value_label])
        working_df = df.copy()
        working_df["_summary_line"] = build_summary_label(
            working_df, summary_column, summary_fallback_column, category_label or section
        )
        if category_column is None:
            working_df["_category_label"] = category_label or section
            grp_cols = ["_category_label", "_summary_line"]
            rename_src = "_category_label"
        else:
            grp_cols = [category_column, "_summary_line"]
            rename_src = category_column
        aggregated = (
            working_df.groupby(grp_cols, dropna=False)
            .agg(**{value_label: (value_column, "sum"), date_label: (date_column, summarize_line_dates)})
            .reset_index()
            .rename(columns={rename_src: "Category", "_summary_line": "Line Item"})
        )
        aggregated.insert(0, "Section", section)
        aggregated["Category"] = aggregated["Category"].fillna(section)
        aggregated["Line Item"] = aggregated["Line Item"].fillna("").astype(str).str.strip()
        return aggregated

    planned_lines = pd.concat([
        aggregate_lines(plan_receivables, section="Collections", category_column=None,
                        summary_column="counterparty", summary_fallback_column="description",
                        date_column="forecast_collection_date", value_column="expected_base",
                        value_label="Planned", date_label="Planned Date", category_label="Inflows"),
        aggregate_lines(plan_outflows, section="Outflows", category_column="sheet_name",
                        summary_column="vendor_name", summary_fallback_column="description",
                        date_column="date", value_column="amount",
                        value_label="Planned", date_label="Planned Date"),
    ], ignore_index=True)

    actual_lines = pd.concat([
        aggregate_lines(actual_receivables, section="Collections", category_column=None,
                        summary_column="counterparty", summary_fallback_column="description",
                        date_column="comparison_date", value_column="comparison_amount",
                        value_label="Actual", date_label="Actual Date", category_label="Inflows"),
        aggregate_lines(actual_outflows, section="Outflows", category_column="sheet_name",
                        summary_column="vendor_name", summary_fallback_column="description",
                        date_column="date", value_column="amount",
                        value_label="Actual", date_label="Actual Date"),
    ], ignore_index=True)

    merged = planned_lines.merge(actual_lines, on=["Section", "Category", "Line Item"], how="outer")
    if merged.empty:
        return pd.DataFrame(columns=columns)

    merged["Planned"] = merged["Planned"].fillna(0.0)
    merged["Actual"] = merged["Actual"].fillna(0.0)
    merged["Planned Date"] = merged["Planned Date"].fillna("")
    merged["Actual Date"] = merged["Actual Date"].fillna("")
    merged["Variance"] = merged["Actual"] - merged["Planned"]
    merged["Variance %"] = np.where(
        merged["Planned"].abs() > 0,
        merged["Variance"].abs() / merged["Planned"].abs(),
        np.nan,
    )
    merged["Status"] = np.select(
        [
            (merged["Planned"] > 0) & (merged["Actual"] > 0),
            (merged["Planned"] > 0) & (merged["Actual"] <= 0),
            (merged["Planned"] <= 0) & (merged["Actual"] > 0),
        ],
        ["Matched", "Planned only", "Actual only"],
        default="Review",
    )
    merged["_abs_variance"] = merged["Variance"].abs()
    merged = merged.sort_values(
        by=["Section", "Category", "_abs_variance", "Line Item"],
        ascending=[True, True, False, True],
    ).drop(columns="_abs_variance")
    return merged[columns].reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
#  FORMATTERS
# ══════════════════════════════════════════════════════════════════════════════

def format_currency(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return f"{value:,.0f}"


def format_percent(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return f"{value:.0%}"


def format_runway(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if value == np.inf:
        return "No burn"
    return f"{value:,.0f} days"


def negative_red(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return "color: #b91c1c;" if float(value) < 0 else ""


def round_numeric_columns(df: pd.DataFrame, percent_columns: list[str] | None = None) -> pd.DataFrame:
    rounded = df.copy()
    percent_set = set(percent_columns or [])
    for column in rounded.columns:
        if column in percent_set or not pd.api.types.is_numeric_dtype(rounded[column]):
            continue
        rounded[column] = rounded[column].apply(
            lambda v: int(round(float(v))) if pd.notna(v) else v
        )
    return rounded


def apply_excel_number_formats(worksheet) -> None:
    percent_headers = {"Variance %"}
    for column_cells in worksheet.columns:
        header = str(column_cells[0].value or "")
        if not header:
            continue
        is_percent_column = ("Probability" in header) or (header in percent_headers)
        for cell in column_cells[1:]:
            if isinstance(cell.value, (datetime, date, pd.Timestamp)):
                cell.number_format = "DD-MM-YYYY"
                continue
            if cell.value is None or isinstance(cell.value, bool) or not isinstance(cell.value, Number):
                continue
            cell.number_format = "0%" if is_percent_column else "#,##0"


def bold_totals(row: pd.Series) -> list[str]:
    line_item = str(row.get("Line Item", "")).strip()
    return [
        "font-weight: 700;" + ("background-color: #fef2f2;" if line_item == "OD Utilization" else "")
        if line_item in TOTAL_ROWS
        else ("background-color: #fef2f2;" if line_item == "OD Utilization" else "")
        for _ in row
    ]


# ══════════════════════════════════════════════════════════════════════════════
#  SORT CONTROLS
# ══════════════════════════════════════════════════════════════════════════════

def apply_sort_controls(
    df: pd.DataFrame,
    *,
    key_prefix: str,
    sortable_columns: list[str] | None = None,
    default_sort_by: str | None = None,
    default_descending: bool = False,
) -> pd.DataFrame:
    if df.empty or len(df) <= 1:
        return df
    available_columns = [c for c in (sortable_columns or df.columns.tolist()) if c in df.columns]
    if not available_columns:
        return df
    default_label = "Default order"
    sort_options = [default_label] + available_columns
    default_index = sort_options.index(default_sort_by) if default_sort_by in available_columns else 0
    c1, c2 = st.columns([1.35, 1])
    selected_column = c1.selectbox("Sort by", sort_options, index=default_index, key=f"{key_prefix}_sort_by")
    selected_order = c2.selectbox(
        "Order", ["Ascending", "Descending"],
        index=1 if default_descending else 0,
        key=f"{key_prefix}_sort_order",
        disabled=selected_column == default_label,
    )
    if selected_column == default_label:
        return df.reset_index(drop=True)
    return df.sort_values(
        by=selected_column, ascending=(selected_order == "Ascending"), na_position="last", kind="stable"
    ).reset_index(drop=True)


def sort_financial_matrix(
    df: pd.DataFrame,
    *,
    key_prefix: str,
    sortable_columns: list[str] | None = None,
    preserve_total_rows: bool = True,
) -> pd.DataFrame:
    if df.empty or len(df) <= 1:
        return df
    available_columns = [c for c in (sortable_columns or df.columns.tolist()) if c in df.columns]
    if not available_columns:
        return df
    default_label = "Default order"
    sort_options = [default_label] + available_columns
    c1, c2 = st.columns([1.35, 1])
    selected_column = c1.selectbox("Sort by", sort_options, index=0, key=f"{key_prefix}_sort_by")
    selected_order = c2.selectbox(
        "Order", ["Ascending", "Descending"], index=0, key=f"{key_prefix}_sort_order",
        disabled=selected_column == default_label,
    )
    if selected_column == default_label:
        return df.reset_index(drop=True)
    totals_mask = (
        df["Line Item"].astype(str).str.strip().isin(TOTAL_ROWS)
        if preserve_total_rows and "Line Item" in df.columns
        else pd.Series(False, index=df.index)
    )
    pinned = df.loc[totals_mask].copy()
    sortable = df.loc[~totals_mask].copy()
    if sortable.empty:
        return df.reset_index(drop=True)
    sorted_rows = sortable.sort_values(
        by=selected_column, ascending=(selected_order == "Ascending"), na_position="last", kind="stable"
    )
    return (sorted_rows if pinned.empty else pd.concat([pinned, sorted_rows], ignore_index=True))


# ══════════════════════════════════════════════════════════════════════════════
#  TABLE RENDERERS
# ══════════════════════════════════════════════════════════════════════════════

def render_financial_table(
    df: pd.DataFrame,
    numeric_columns: list[str],
    height: int | None = None,
    *,
    enable_sort: bool = False,
    sort_key_prefix: str | None = None,
    sortable_columns: list[str] | None = None,
    preserve_total_rows: bool = True,
) -> None:
    if df.empty:
        st.info("No records available for this view.")
        return
    display_df = df.copy()
    if enable_sort and sort_key_prefix:
        display_df = sort_financial_matrix(
            display_df, key_prefix=sort_key_prefix,
            sortable_columns=sortable_columns, preserve_total_rows=preserve_total_rows,
        )
    display_df = format_date_columns(display_df)
    styler = (
        display_df.style.hide(axis="index")
        .format({col: format_currency for col in numeric_columns}, na_rep="")
        .apply(bold_totals, axis=1)
        .applymap(negative_red, subset=numeric_columns)
        .set_properties(subset=["Line Item"], **{"text-align": "left", "white-space": "pre"})
        .set_properties(subset=numeric_columns, **{"text-align": "right"})
        .set_table_styles([
            {"selector": "table", "props": [("width", "100%"), ("border-collapse", "collapse"), ("font-size", "0.95rem")]},
            {"selector": "th", "props": [("text-align", "right"), ("padding", "8px 10px"), ("border-bottom", "1px solid #d4d4d8")]},
            {"selector": "th.col_heading.level0.col0", "props": [("text-align", "left")]},
            {"selector": "td", "props": [("padding", "8px 10px"), ("border-bottom", "1px solid #ececf1")]},
        ])
    )
    st.markdown(styler.to_html(), unsafe_allow_html=True)


def render_simple_table(
    df: pd.DataFrame,
    currency_columns: list[str] | None = None,
    percent_columns: list[str] | None = None,
    *,
    enable_sort: bool = False,
    sort_key_prefix: str | None = None,
    sortable_columns: list[str] | None = None,
    default_sort_by: str | None = None,
    default_descending: bool = False,
) -> None:
    if df.empty:
        st.info("No records available.")
        return
    currency_columns = currency_columns or []
    percent_columns = percent_columns or []
    display_df = df.copy()
    if enable_sort and sort_key_prefix:
        display_df = apply_sort_controls(
            display_df, key_prefix=sort_key_prefix, sortable_columns=sortable_columns,
            default_sort_by=default_sort_by, default_descending=default_descending,
        )
    display_df = format_date_columns(display_df)
    styler = display_df.style.hide(axis="index")
    formatters: dict[str, Any] = {}
    if currency_columns:
        formatters.update({col: format_currency for col in currency_columns})
        styler = styler.applymap(negative_red, subset=currency_columns)
    if percent_columns:
        formatters.update({col: format_percent for col in percent_columns})
    if formatters:
        styler = styler.format(formatters, na_rep="")
    styler = styler.set_table_styles([
        {"selector": "table", "props": [("width", "max-content"), ("min-width", "100%"), ("border-collapse", "collapse"), ("font-size", "0.94rem")]},
        {"selector": "th", "props": [("text-align", "left"), ("padding", "8px 10px"), ("border-bottom", "1px solid #d4d4d8"), ("vertical-align", "top"), ("white-space", "nowrap")]},
        {"selector": "td", "props": [("padding", "8px 10px"), ("border-bottom", "1px solid #ececf1"), ("vertical-align", "top"), ("white-space", "normal"), ("word-break", "break-word")]},
    ])
    st.markdown(
        f'<div style="width:100%; overflow-x:auto;">{styler.to_html()}</div>',
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
#  CFO ALERT BANNER  (new)
# ══════════════════════════════════════════════════════════════════════════════

def render_cash_alert_banner(
    projected_ending: float,
    min_cash_threshold: float,
    od_utilization: float,
    od_sanctioned_limit: float,
    od_interest_est: float,
) -> None:
    """
    Display a red/amber alert strip when cash < threshold or OD is at limit.
    Only visible when there is an active risk condition.
    """
    alerts: list[tuple[str, str]] = []   # (severity, message)

    if min_cash_threshold > 0 and projected_ending < min_cash_threshold:
        if projected_ending < 0:
            alerts.append(("red", f"CRITICAL: Projected month-end cash is negative ({format_currency(projected_ending)}). Immediate action required."))
        else:
            gap = min_cash_threshold - projected_ending
            alerts.append(("amber", f"WARNING: Projected closing cash ({format_currency(projected_ending)}) is below the minimum floor of {format_currency(min_cash_threshold)} — shortfall of {format_currency(gap)}."))

    if od_sanctioned_limit > 0 and od_utilization >= 0.90 * od_sanctioned_limit:
        util_pct = od_utilization / od_sanctioned_limit * 100
        alerts.append(("amber", f"OD HEADROOM TIGHT: Utilising {util_pct:.0f}% of sanctioned limit ({format_currency(od_utilization)} of {format_currency(od_sanctioned_limit)})."))

    if od_interest_est > 0:
        alerts.append(("info", f"Estimated OD interest for this month: {format_currency(od_interest_est)} (based on {od_interest_est:.0f} drawn at configured rate)."))

    for severity, message in alerts:
        colour_map = {
            "red": ("#fef2f2", "#991b1b", "#fecaca"),
            "amber": ("#fffbeb", "#92400e", "#fde68a"),
            "info": ("#eff6ff", "#1e40af", "#bfdbfe"),
        }
        bg, text, border = colour_map.get(severity, colour_map["info"])
        st.markdown(
            f"""
            <div style="border:1px solid {border}; border-left:4px solid {text};
                        background:{bg}; color:{text}; padding:12px 16px;
                        margin:8px 0; border-radius:4px; font-size:0.94rem; font-weight:600;">
                {message}
            </div>
            """,
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
#  KPI SECTION
# ══════════════════════════════════════════════════════════════════════════════

def render_kpis(
    opening_balance: float,
    weekly_cashflow: dict[str, Any],
    scenario_table: pd.DataFrame,
    next_30_day_requirement: float,
    overdue_dependency: float,
    sixty_plus_dependency: float,
    *,
    min_cash_threshold: float = 0.0,
    od_sanctioned_limit: float = 0.0,
) -> None:
    plan = weekly_cashflow["selected_month_plan"]
    base_case = scenario_table.loc[scenario_table["Scenario"] == "Base Case"].iloc[0] if not scenario_table.empty else None
    runway_days = base_case["Runway"] if base_case is not None else np.inf

    # ── Alert banners ────────────────────────────────────────────────────────
    render_cash_alert_banner(
        projected_ending=plan["ending"],
        min_cash_threshold=min_cash_threshold,
        od_utilization=plan.get("od_utilization", 0.0),
        od_sanctioned_limit=od_sanctioned_limit,
        od_interest_est=plan.get("od_interest_est", 0.0),
    )

    # ── Primary KPI row ──────────────────────────────────────────────────────
    cols = st.columns(5)
    cols[0].metric("Opening Cash", format_currency(opening_balance))
    cols[1].metric("Net Inflows (Month)", format_currency(plan["net"]))
    cols[2].metric("Total Outflows (Month)", format_currency(plan["outflows"]))
    cols[3].metric("Month-End Cash (Planned)", format_currency(plan["ending"]))
    cols[4].metric("Runway", format_runway(runway_days))

    # ── OD / credit-line KPI row ─────────────────────────────────────────────
    od_util = plan.get("od_utilization", 0.0) or 0.0
    if od_sanctioned_limit > 0 or od_util > 0:
        od_headroom = max(od_sanctioned_limit - od_util, 0.0)
        od_pct = (od_util / od_sanctioned_limit * 100) if od_sanctioned_limit > 0 else 0.0
        od_interest = plan.get("od_interest_est", 0.0) or 0.0
        od_cols = st.columns(4)
        od_cols[0].metric("OD Sanctioned Limit", format_currency(od_sanctioned_limit) if od_sanctioned_limit > 0 else "—")
        od_cols[1].metric("OD Utilization (Est.)", format_currency(od_util))
        od_cols[2].metric("OD Headroom", format_currency(od_headroom))
        od_cols[3].metric("Est. OD Interest Cost", format_currency(od_interest) if od_interest > 0 else "Nil")

    # ── Decision box ─────────────────────────────────────────────────────────
    statement, actions = build_decision_box(
        runway_days, overdue_dependency, sixty_plus_dependency, plan["ending"],
        min_cash_threshold=min_cash_threshold,
    )
    bullets = "".join(f"<li>{action}</li>" for action in actions[:3])
    st.markdown(
        f"""
        <div style="border:1px solid #d4d4d8; border-left:4px solid #111827;
                    padding:16px 18px; margin-top:10px; margin-bottom:18px; background:#ffffff;">
            <div style="font-size:1.0rem; font-weight:700; margin-bottom:8px;">Decision Box</div>
            <div style="font-size:0.98rem; margin-bottom:10px;">{statement}</div>
            <ul style="margin:0; padding-left:18px;">{bullets}</ul>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_decision_box(
    runway_days: float,
    overdue_dependency: float,
    sixty_plus_dependency: float,
    projected_closing_cash: float,
    *,
    min_cash_threshold: float = 0.0,
) -> tuple[str, list[str]]:
    if runway_days != np.inf and runway_days <= 30:
        sentence = f"Cash will run out in approximately {max(runway_days, 0):.0f} days. Primary risk is near-term liquidity compression."
    elif projected_closing_cash < 0:
        sentence = "Projected month-end cash is negative. Primary risk is a forecast shortfall against committed payments."
    elif min_cash_threshold > 0 and projected_closing_cash < min_cash_threshold:
        sentence = f"Projected closing cash ({format_currency(projected_closing_cash)}) is below the minimum operating floor ({format_currency(min_cash_threshold)}). Buffer must be rebuilt before discretionary payments."
    elif overdue_dependency > 0.30:
        sentence = "Liquidity is dependent on overdue collections. Primary risk is slippage in already aged receivables."
    else:
        sentence = "Liquidity is currently manageable. Primary watchpoint is preserving forecast collections against committed outflows."

    actions: list[str] = []
    if overdue_dependency > 0.30:
        actions.append("Escalate overdue collections immediately and track owner-level recovery actions.")
    if sixty_plus_dependency > 0.30:
        actions.append("Stress-test the plan assuming 60+ day receivables slip beyond the month.")
    if projected_closing_cash < 0 or (runway_days != np.inf and runway_days <= 30):
        actions.append("Sequence discretionary payments after payroll, statutory dues, and critical vendors.")
    if min_cash_threshold > 0 and projected_closing_cash < min_cash_threshold:
        actions.append(f"Preserve minimum cash floor of {format_currency(min_cash_threshold)}. Review deferred or discretionary payments first.")
    if not actions:
        actions.append("Monitor weekly receipts against the base case and refresh the forecast at each weekly close.")
    return sentence, actions[:3]


# ══════════════════════════════════════════════════════════════════════════════
#  CFO COMMENTARY BOX  (new)
# ══════════════════════════════════════════════════════════════════════════════

def render_cfo_commentary(section_key: str, label: str = "CFO Commentary") -> None:
    """
    Persistent, editable commentary box. Content is stored in session state
    and included in Excel exports.
    """
    st.markdown(f"**{label}**")
    st.text_area(
        label,
        key=f"commentary_{section_key}",
        placeholder="Add narrative, key decisions, or variance explanations here. This will be included in the Excel export.",
        height=90,
        label_visibility="collapsed",
    )


# ══════════════════════════════════════════════════════════════════════════════
#  SECTION RENDERERS
# ══════════════════════════════════════════════════════════════════════════════

def render_collection_engine(
    scenario_table: pd.DataFrame,
    overdue_dependency: float,
    sixty_plus_dependency: float,
) -> None:
    scenario_display = scenario_table.copy()
    if not scenario_display.empty:
        scenario_display["Runway"] = scenario_display["Runway"].map(format_runway)
    render_simple_table(scenario_display, currency_columns=["Expected Collection", "Impact on Closing Cash"])

    flag_text = (
        "Flag: More than 30% of forecast collections are coming from 60+ day receivables."
        if sixty_plus_dependency > 0.30
        else "Flag: 60+ day receivable dependency remains within the 30% guardrail."
    )
    st.markdown(
        f"""
        <div style="border:1px solid #d4d4d8; padding:14px 16px; margin-top:12px; background:#ffffff;">
            <div><strong>% dependency on overdue collections:</strong> {format_percent(overdue_dependency)}</div>
            <div style="margin-top:6px; color:{'#b91c1c' if sixty_plus_dependency > 0.30 else '#111827'};">{flag_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_six_month_view(
    monthly_forecast: pd.DataFrame,
    monthly_inflow_detail: pd.DataFrame,
    monthly_outflow_detail: pd.DataFrame,
    *,
    min_cash_threshold: float = 0.0,
) -> None:
    if monthly_forecast.empty:
        st.info("No monthly forecast available.")
        return

    display = monthly_forecast[["Month", "Opening Cash", "Inflows", "Outflows", "Closing Cash"]].copy()

    def style_closing_row(row: pd.Series) -> list[str]:
        closing = row["Closing Cash"]
        base = "font-weight: 700; "
        if closing < 0:
            return [base + "color: #b91c1c;" for _ in row]
        if min_cash_threshold > 0 and closing < min_cash_threshold:
            return [base + "color: #92400e;" for _ in row]
        return ["" for _ in row]

    styler = (
        display.style.hide(axis="index")
        .format({
            "Opening Cash": format_currency,
            "Inflows": format_currency,
            "Outflows": format_currency,
            "Closing Cash": format_currency,
        }, na_rep="")
        .applymap(negative_red, subset=["Closing Cash"])
        .apply(style_closing_row, axis=1)
        .set_table_styles([
            {"selector": "table", "props": [("width", "100%"), ("border-collapse", "collapse"), ("font-size", "0.95rem")]},
            {"selector": "th", "props": [("text-align", "left"), ("padding", "8px 10px"), ("border-bottom", "1px solid #d4d4d8")]},
            {"selector": "td", "props": [("padding", "8px 10px"), ("border-bottom", "1px solid #ececf1")]},
        ])
    )
    st.markdown(styler.to_html(), unsafe_allow_html=True)

    if min_cash_threshold > 0:
        breach_months = monthly_forecast.loc[monthly_forecast["Closing Cash"] < min_cash_threshold, "Month"].tolist()
        if breach_months:
            st.markdown(
                f"""
                <div style="border:1px solid #fde68a; border-left:4px solid #92400e;
                            background:#fffbeb; color:#92400e; padding:10px 14px;
                            margin:8px 0; border-radius:4px; font-size:0.9rem;">
                    <strong>Cash floor breach:</strong> {', '.join(breach_months)} — closing cash below minimum threshold of {format_currency(min_cash_threshold)}.
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("**Monthly Inflow Detail**")
    render_financial_table(
        monthly_inflow_detail,
        numeric_columns=[c for c in monthly_inflow_detail.columns if c != "Line Item"],
        enable_sort=True,
        sort_key_prefix="six_month_inflow_detail",
        sortable_columns=monthly_inflow_detail.columns.tolist(),
    )
    st.markdown("**Monthly Expense Category Detail**")
    render_financial_table(
        monthly_outflow_detail,
        numeric_columns=[c for c in monthly_outflow_detail.columns if c != "Line Item"],
        enable_sort=True,
        sort_key_prefix="six_month_outflow_detail",
        sortable_columns=monthly_outflow_detail.columns.tolist(),
    )


def render_actual_vs_budget(
    plan_results: dict[str, Any],
    actual_data: dict[str, Any] | None,
    actual_file_name: str | None,
    *,
    base_probs: dict[str, float] | None = None,
    worst_probs: dict[str, float] | None = None,
    tds_rate: float = 0.0,
) -> None:
    comparison = compare_actual_vs_budget(plan_results, actual_data)
    line_variance = build_line_level_variance(
        plan_results, actual_data,
        base_probs=base_probs, worst_probs=worst_probs, tds_rate=tds_rate,
    )
    left_col, right_col = st.columns([1, 2])

    with left_col:
        st.markdown("**Actual Data**")
        if actual_file_name:
            st.success(f"Actual workbook loaded: {actual_file_name}")
            st.caption("Use the sidebar uploader to replace or clear the actual workbook.")
        else:
            st.info("Upload an actuals workbook from the sidebar to unlock variance tracking.")

    with right_col:
        st.markdown("**Comparison**")
        if comparison.empty:
            st.info("Actual data not available for comparison.")
        else:
            comparison = apply_sort_controls(
                comparison, key_prefix="actual_vs_budget_comparison",
                sortable_columns=["Line Item", "Planned", "Actual", "Variance", "Variance %"],
            )
            major_deviation_mask = comparison["Variance %"].fillna(0) > 0.10
            cash_impact = comparison.loc[comparison["Line Item"] == "Ending Cash Balance", "Variance"].sum()
            styler = (
                comparison.style.hide(axis="index")
                .format({"Planned": format_currency, "Actual": format_currency, "Variance": format_currency, "Variance %": format_percent}, na_rep="")
                .applymap(negative_red, subset=["Variance"])
                .apply(
                    lambda row: [
                        "background-color: #fef2f2;"
                        if (row["Variance %"] if not pd.isna(row["Variance %"]) else 0) > 0.10
                        else ""
                        for _ in row
                    ],
                    axis=1,
                )
                .set_table_styles([
                    {"selector": "table", "props": [("width", "100%"), ("border-collapse", "collapse"), ("font-size", "0.95rem")]},
                    {"selector": "th", "props": [("text-align", "left"), ("padding", "8px 10px"), ("border-bottom", "1px solid #d4d4d8")]},
                    {"selector": "td", "props": [("padding", "8px 10px"), ("border-bottom", "1px solid #ececf1")]},
                ])
            )
            st.markdown(styler.to_html(), unsafe_allow_html=True)
            st.caption(f"Cash impact of variance: {format_currency(cash_impact)}")

            major_deviations = comparison.loc[major_deviation_mask, ["Line Item", "Variance", "Variance %"]]
            if not major_deviations.empty:
                st.caption("Major deviations (>10%)")
                render_simple_table(
                    major_deviations,
                    currency_columns=["Variance"],
                    percent_columns=["Variance %"],
                    enable_sort=True,
                    sort_key_prefix="actual_vs_budget_major_deviations",
                    sortable_columns=["Line Item", "Variance", "Variance %"],
                )

    # ── Variance Commentary ────────────────────────────────────────────────────
    if not comparison.empty:
        render_cfo_commentary("avb_variance", label="Variance Commentary")

    if comparison.empty:
        return

    st.markdown("**Summary-Level Variance**")
    if line_variance.empty:
        st.info("No selected-month summary rows available for comparison.")
        return

    f1, f2, f3 = st.columns([1, 1, 1.4])
    section_options = ["All"] + sorted(line_variance["Section"].dropna().astype(str).unique().tolist())
    selected_section = f1.selectbox("Section", section_options, key="avb_line_section")
    cat_src = line_variance if selected_section == "All" else line_variance.loc[line_variance["Section"] == selected_section]
    cat_options = ["All"] + sorted(cat_src["Category"].dropna().astype(str).unique().tolist())
    selected_category = f2.selectbox("Category", cat_options, key="avb_line_category")
    li_src = cat_src if selected_category == "All" else cat_src.loc[cat_src["Category"] == selected_category]
    selected_line_items = f3.multiselect(
        "Summary Lines",
        options=sorted(li_src["Line Item"].dropna().astype(str).unique().tolist()),
        placeholder="All vendors or clients",
        key="avb_line_items",
    )

    c1, c2, c3, c4 = st.columns([1, 1, 0.9, 1.2])
    status_options = ["All"] + sorted(line_variance["Status"].dropna().astype(str).unique().tolist())
    selected_status = c1.selectbox("Status", status_options, key="avb_status")
    sort_order = c2.selectbox(
        "Sort By",
        ["Largest Variance", "Largest Variance %", "Largest Planned", "Largest Actual", "A-Z"],
        key="avb_sort_order",
    )
    only_exceptions = c3.checkbox("Only Exceptions", value=False, key="avb_only_exceptions")
    search_text = c4.text_input("Search", value="", placeholder="Vendor, client, or category", key="avb_search").strip().lower()

    flv = line_variance.copy()
    if selected_section != "All":
        flv = flv.loc[flv["Section"] == selected_section]
    if selected_category != "All":
        flv = flv.loc[flv["Category"] == selected_category]
    if selected_line_items:
        flv = flv.loc[flv["Line Item"].isin(selected_line_items)]
    if selected_status != "All":
        flv = flv.loc[flv["Status"] == selected_status]
    if only_exceptions:
        flv = flv.loc[(flv["Status"] != "Matched") | (flv["Variance %"].fillna(0) > 0.10)]
    if search_text:
        flv = flv.loc[
            flv["Category"].astype(str).str.lower().str.contains(search_text, na=False)
            | flv["Line Item"].astype(str).str.lower().str.contains(search_text, na=False)
        ]

    sort_map = {
        "Largest Variance": lambda df: df.assign(_s=df["Variance"].abs()).sort_values(["_s", "Line Item"], ascending=[False, True]),
        "Largest Variance %": lambda df: df.assign(_s=df["Variance %"].fillna(-1)).sort_values(["_s", "Line Item"], ascending=[False, True]),
        "Largest Planned": lambda df: df.sort_values(["Planned", "Line Item"], ascending=[False, True]),
        "Largest Actual": lambda df: df.sort_values(["Actual", "Line Item"], ascending=[False, True]),
        "A-Z": lambda df: df.sort_values(["Section", "Category", "Line Item"]),
    }
    flv = sort_map.get(sort_order, lambda df: df)(flv)
    if "_s" in flv.columns:
        flv = flv.drop(columns="_s")

    if flv.empty:
        st.info("No line items match the selected filters.")
        return

    status_summary = flv["Status"].value_counts()
    m1, m2, m3 = st.columns(3)
    m1.metric("Matched Summary Rows", int(status_summary.get("Matched", 0)))
    m2.metric("Planned Only", int(status_summary.get("Planned only", 0)))
    m3.metric("Actual Only", int(status_summary.get("Actual only", 0)))
    st.caption(f"Selected-month variance · Rows shown: {len(flv):,}")
    render_simple_table(flv, currency_columns=["Planned", "Actual", "Variance"], percent_columns=["Variance %"])

    line_exceptions = flv.loc[
        (flv["Status"] != "Matched") | (flv["Variance %"].fillna(0) > 0.10),
        ["Section", "Category", "Line Item", "Status", "Variance", "Variance %"],
    ]
    if not line_exceptions.empty:
        st.caption("Summary-level exceptions: unmatched or >10% variance")
        render_simple_table(
            line_exceptions,
            currency_columns=["Variance"],
            percent_columns=["Variance %"],
            enable_sort=True,
            sort_key_prefix="avb_line_exceptions",
            sortable_columns=["Section", "Category", "Line Item", "Status", "Variance", "Variance %"],
        )


def render_weekly_drilldown(week_meta: list[dict[str, Any]]) -> None:
    st.markdown("**Weekly Drill-Down**")
    for week in week_meta:
        with st.expander(week["label"], expanded=False):
            dl_col, _ = st.columns([1, 3])
            with dl_col:
                week_export = build_week_raw_export(week)
                slug = week["key"].lower().replace(" ", "_")
                st.download_button(
                    "Download Raw Week Data (Excel)",
                    data=week_export,
                    file_name=f"{slug}_raw_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{slug}",
                    use_container_width=True,
                )
            if week.get("holidays"):
                holiday_text = "; ".join(
                    f"{full_date_with_day(e['date'])}: {e['name']}" for e in week["holidays"]
                )
                st.markdown(
                    f"""
                    <div style="border:1px solid #fecaca; background:#fef2f2; color:#991b1b;
                                padding:10px 12px; margin-bottom:12px; border-radius:6px;">
                        <strong>Anticipated bank holidays:</strong> {holiday_text}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            inflows = summarize_inflow_detail(week["plan_receipts_detail"], value_column="expected_base", value_label="Expected Collection")
            outflows = summarize_outflow_detail(week["plan_outflows_detail"])
            detail_filter = st.radio(
                "Weekly detail view", options=["Inflows", "Outflows"], index=0,
                horizontal=True,
                key=f"weekly_detail_{week['key'].lower().replace(' ', '_')}",
                label_visibility="collapsed",
            )
            slug = week["key"].lower().replace(" ", "_")
            if detail_filter == "Inflows":
                st.caption("Summary inflows grouped by billing type and client.")
                render_simple_table(inflows, currency_columns=["Expected Collection"], enable_sort=True,
                                    sort_key_prefix=f"{slug}_inflows",
                                    sortable_columns=["Billing Type", "Client", "Due Window", "Forecast Window", "Expected Collection", "Source Rows"])
            else:
                st.caption("Summary outflows grouped by category and vendor.")
                render_simple_table(outflows, currency_columns=["Amount"], enable_sort=True,
                                    sort_key_prefix=f"{slug}_outflows",
                                    sortable_columns=["Category", "Vendor", "Payment Window", "Amount", "Source Rows"])


def render_week_holiday_summary(week_meta: list[dict[str, Any]]) -> None:
    st.caption(
        "Bank-holiday overlay covers the selected region. Second and fourth Saturdays are included. Sundays excluded."
    )
    columns = st.columns(max(len(week_meta), 1))
    for col, week in zip(columns, week_meta):
        with col:
            with st.container(border=True):
                st.markdown(f"**{week['key']}**")
                if week.get("holidays"):
                    holiday_lines = "\n".join(
                        f"- {short_date_label(e['date'])}: {e['name']}" for e in week["holidays"]
                    )
                    st.markdown(holiday_lines)
                else:
                    st.caption("No anticipated bank holidays")


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def _auto_column_widths(sheet) -> None:
    for col_cells in sheet.columns:
        max_length = max((len(str(cell.value or "")) for cell in col_cells), default=0)
        sheet.column_dimensions[col_cells[0].column_letter].width = min(max_length + 2, 42)


def _get_commentary(key: str) -> str:
    return st.session_state.get(f"commentary_{key}", "")


def build_weekly_breakdown_export(
    weekly_cashflow: dict[str, Any],
    selected_month: pd.Timestamp,
) -> bytes:
    output = BytesIO()
    weekly_table = round_numeric_columns(weekly_cashflow["weekly_table"].copy())

    holiday_rows = []
    for week in weekly_cashflow["week_meta"]:
        if week.get("holidays"):
            for e in week["holidays"]:
                holiday_rows.append({"Week": week["key"], "Week Range": week["label"],
                                     "Holiday Date": format_date(e["date"]), "Holiday Name": e["name"]})
        else:
            holiday_rows.append({"Week": week["key"], "Week Range": week["label"],
                                 "Holiday Date": "", "Holiday Name": "No anticipated bank holidays"})
    holiday_df = pd.DataFrame(holiday_rows)

    # Commentary
    exec_commentary = _get_commentary("executive_summary")
    avb_commentary = _get_commentary("avb_variance")
    commentary_df = pd.DataFrame([
        {"Section": "Executive Summary", "Commentary": exec_commentary},
        {"Section": "Actual vs Budget", "Commentary": avb_commentary},
    ])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        weekly_table.to_excel(writer, index=False, sheet_name="Weekly Breakdown")
        format_date_columns(holiday_df).to_excel(writer, index=False, sheet_name="Week Holidays")
        commentary_df.to_excel(writer, index=False, sheet_name="CFO Commentary")

        ws = writer.sheets["Weekly Breakdown"]
        ws.freeze_panes = "B2"
        apply_excel_number_formats(ws)
        _auto_column_widths(ws)
        _auto_column_widths(writer.sheets["Week Holidays"])
        _auto_column_widths(writer.sheets["CFO Commentary"])

    output.seek(0)
    return output.getvalue()


def build_week_raw_export(week: dict[str, Any]) -> bytes:
    output = BytesIO()
    inflows = round_numeric_columns(
        week["plan_receipts_detail"].copy(),
        percent_columns=["base_probability", "best_probability", "worst_probability"],
    )
    outflows = round_numeric_columns(week["plan_outflows_detail"].copy())
    inflow_summary = round_numeric_columns(
        summarize_inflow_detail(week["plan_receipts_detail"], value_column="expected_base", value_label="Expected Collection")
    )
    outflow_summary = round_numeric_columns(summarize_outflow_detail(week["plan_outflows_detail"]))

    if inflows.empty:
        inflows = pd.DataFrame(columns=["counterparty", "description", "billing_type", "billing_bucket",
                                        "due_date", "aging_bucket", "base_probability", "expected_base", "forecast_collection_date"])
    else:
        keep_cols = ["invoice_date", "tentative_collection_date", "forecast_collection_date", "due_date",
                     "billing_type", "billing_bucket", "counterparty", "description", "amount", "aging_bucket",
                     "base_probability", "best_probability", "worst_probability",
                     "expected_base", "expected_best", "expected_worst"]
        inflows = inflows[[c for c in keep_cols if c in inflows.columns]].rename(columns={
            "invoice_date": "Invoice Date", "tentative_collection_date": "Tentative Collection Date",
            "forecast_collection_date": "Forecast Collection Date", "due_date": "Due Date",
            "billing_type": "Billing Type", "billing_bucket": "Billing Bucket",
            "counterparty": "Client", "description": "Reference", "amount": "Invoice Amount",
            "aging_bucket": "Aging Bucket", "base_probability": "Base Probability",
            "best_probability": "Best Probability", "worst_probability": "Worst Probability",
            "expected_base": "Expected Base Collection", "expected_best": "Expected Best Collection",
            "expected_worst": "Expected Worst Collection",
        })

    if outflows.empty:
        outflows = pd.DataFrame(columns=["sheet_name", "vendor_name", "description", "date", "amount"])
    else:
        outflows = outflows[["sheet_name", "vendor_name", "description", "date", "amount"]].rename(columns={
            "sheet_name": "Category", "vendor_name": "Vendor",
            "description": "Particular", "date": "Payment Date", "amount": "Amount",
        })

    holiday_rows = [
        {"Week": week["key"], "Week Range": week["label"],
         "Holiday Date": format_date(e["date"]), "Holiday Name": e["name"]}
        for e in week.get("holidays", [])
    ]
    holiday_df = (
        pd.DataFrame(holiday_rows)
        if holiday_rows
        else pd.DataFrame([{"Week": week["key"], "Week Range": week["label"],
                            "Holiday Date": "", "Holiday Name": "No anticipated bank holidays"}])
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        inflow_summary.to_excel(writer, index=False, sheet_name="Inflows Summary")
        outflow_summary.to_excel(writer, index=False, sheet_name="Outflows Summary")
        inflows.to_excel(writer, index=False, sheet_name="Inflows Raw")
        outflows.to_excel(writer, index=False, sheet_name="Outflows Raw")
        format_date_columns(holiday_df).to_excel(writer, index=False, sheet_name="Holiday Overlay")
        for sheet in writer.sheets.values():
            sheet.freeze_panes = "A2"
            apply_excel_number_formats(sheet)
            _auto_column_widths(sheet)

    output.seek(0)
    return output.getvalue()


def build_month_raw_export(weekly_cashflow: dict[str, Any], selected_month: pd.Timestamp) -> bytes:
    output = BytesIO()
    month_start = weekly_cashflow["selected_month_start"]
    month_end = weekly_cashflow["selected_month_end"]
    period_start = weekly_cashflow.get("selected_period_start", month_start)

    inflow_frames, outflow_frames, holiday_rows = [], [], []
    for week in weekly_cashflow["week_meta"]:
        if not week["plan_receipts_detail"].empty:
            iw = week["plan_receipts_detail"].copy()
            iw["Week"] = week["key"]
            iw["Week Range"] = week["label"]
            inflow_frames.append(iw)
        if not week["plan_outflows_detail"].empty:
            ow = week["plan_outflows_detail"].copy()
            ow["Week"] = week["key"]
            ow["Week Range"] = week["label"]
            outflow_frames.append(ow)
        if week.get("holidays"):
            for e in week["holidays"]:
                holiday_rows.append({"Week": week["key"], "Week Range": week["label"],
                                     "Holiday Date": format_date(e["date"]), "Holiday Name": e["name"]})

    inflows = pd.concat(inflow_frames, ignore_index=True) if inflow_frames else pd.DataFrame()
    outflows = pd.concat(outflow_frames, ignore_index=True) if outflow_frames else pd.DataFrame()
    inflows = round_numeric_columns(inflows, percent_columns=["base_probability", "best_probability", "worst_probability"]) if not inflows.empty else inflows
    outflows = round_numeric_columns(outflows) if not outflows.empty else outflows

    all_inflows = pd.concat(inflow_frames, ignore_index=True) if inflow_frames else pd.DataFrame()
    all_outflows = pd.concat(outflow_frames, ignore_index=True) if outflow_frames else pd.DataFrame()
    inflow_summary = round_numeric_columns(
        summarize_inflow_detail(all_inflows, value_column="expected_base", value_label="Expected Collection")
    )
    outflow_summary = round_numeric_columns(summarize_outflow_detail(all_outflows))

    keep_inflow_cols = ["Week", "Week Range", "invoice_date", "tentative_collection_date",
                        "forecast_collection_date", "due_date", "billing_type", "billing_bucket",
                        "counterparty", "description", "amount", "aging_bucket",
                        "base_probability", "expected_base"]
    if not inflows.empty:
        inflows = inflows[[c for c in keep_inflow_cols if c in inflows.columns]]
    else:
        inflows = pd.DataFrame(columns=keep_inflow_cols)

    if not outflows.empty:
        outflows = outflows[[c for c in ["Week", "Week Range", "sheet_name", "vendor_name", "description", "date", "amount"] if c in outflows.columns]].rename(columns={
            "sheet_name": "Category", "vendor_name": "Vendor", "description": "Particular", "date": "Payment Date", "amount": "Amount"
        })
    else:
        outflows = pd.DataFrame(columns=["Week", "Week Range", "Category", "Vendor", "Particular", "Payment Date", "Amount"])

    holiday_df = (
        pd.DataFrame(holiday_rows)
        if holiday_rows
        else pd.DataFrame([{"Week": "", "Week Range": "", "Holiday Date": "", "Holiday Name": "No anticipated bank holidays"}])
    )

    cover_df = pd.DataFrame([
        {"Field": "Selected Month", "Value": month_label(selected_month)},
        {"Field": "Balance As Of Date", "Value": format_date(period_start)},
        {"Field": "Month Start", "Value": format_date(month_start)},
        {"Field": "Month End", "Value": format_date(month_end)},
        {"Field": "Weeks in Export", "Value": len(weekly_cashflow["week_meta"])},
        {"Field": "CFO Commentary", "Value": _get_commentary("executive_summary")},
    ])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        format_date_columns(cover_df).to_excel(writer, index=False, sheet_name="Summary")
        inflow_summary.to_excel(writer, index=False, sheet_name="Inflows Summary")
        outflow_summary.to_excel(writer, index=False, sheet_name="Outflows Summary")
        inflows.to_excel(writer, index=False, sheet_name="Inflows Raw")
        outflows.to_excel(writer, index=False, sheet_name="Outflows Raw")
        format_date_columns(holiday_df).to_excel(writer, index=False, sheet_name="Holiday Overlay")
        for sheet in writer.sheets.values():
            sheet.freeze_panes = "A2"
            apply_excel_number_formats(sheet)
            _auto_column_widths(sheet)

    output.seek(0)
    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  STYLES & LAYOUT
# ══════════════════════════════════════════════════════════════════════════════

def inject_styles() -> None:
    st.markdown(
        """
        <style>
        .stApp { background: #f7f7f5; color: #111827; }
        .block-container { padding-top: 1.5rem; padding-bottom: 2.0rem; max-width: 1460px; }
        div[data-testid="metric-container"] {
            border: 1px solid #e5e7eb; padding: 12px 14px; border-radius: 6px;
            background: #ffffff; box-shadow: 0 1px 2px rgba(17,24,39,0.04);
        }
        div[data-testid="metric-container"] label { font-size: 0.82rem; }
        [data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid #e5e7eb; }
        .app-banner {
            background: #ffffff; border: 1px solid #e5e7eb;
            border-left: 4px solid #111827; padding: 18px 20px; margin-bottom: 18px;
        }
        .app-banner-title { font-size: 1.05rem; font-weight: 700; margin-bottom: 6px; }
        .app-banner-subtitle { font-size: 0.93rem; color: #4b5563; }
        .context-grid {
            display: grid; grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 10px; margin-top: 14px;
        }
        .context-card {
            background: #fafaf9; border: 1px solid #e7e5e4;
            padding: 12px 14px; border-radius: 6px;
        }
        .context-label {
            font-size: 0.75rem; color: #6b7280;
            text-transform: uppercase; letter-spacing: 0.03em; margin-bottom: 3px;
        }
        .context-value { font-size: 0.95rem; font-weight: 600; color: #111827; }
        .section-header { margin-top: 18px; margin-bottom: 12px; }
        .section-title { font-size: 1.05rem; font-weight: 700; color: #111827; margin-bottom: 2px; }
        .section-subtitle { font-size: 0.90rem; color: #6b7280; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def default_month_index(available_months: list[pd.Timestamp]) -> int:
    current_month = build_month_start(TODAY)
    for idx, month in enumerate(available_months):
        if build_month_start(month) == current_month:
            return idx
    return 0


def render_section_header(title: str, subtitle: str) -> None:
    st.markdown(
        f"""
        <div class="section-header">
            <div class="section-title">{title}</div>
            <div class="section-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_app_banner(
    selected_month: pd.Timestamp,
    balance_as_of_date: pd.Timestamp,
    planned_source_name: str,
    outflow_categories: int,
    holiday_region: str,
    actual_file_name: str | None,
    load_timestamp: str | None = None,
    tds_rate: float = 0.0,
    min_cash_threshold: float = 0.0,
) -> None:
    actual_status = actual_file_name if actual_file_name else "Not loaded"
    load_ts_display = load_timestamp if load_timestamp else "—"
    tds_display = f"{tds_rate*100:.1f}%" if tds_rate > 0 else "None"
    floor_display = format_currency(min_cash_threshold) if min_cash_threshold > 0 else "—"
    st.markdown(
        f"""
        <div class="app-banner">
            <div class="app-banner-title">Liquidity review workspace for finance leadership</div>
            <div class="app-banner-subtitle">
                Overview first, then weekly cashflow detail, collection scenarios, roll-forward, and actual-vs-budget validation.
            </div>
            <div class="context-grid">
                <div class="context-card">
                    <div class="context-label">Selected Month</div>
                    <div class="context-value">{month_label(selected_month)}</div>
                </div>
                <div class="context-card">
                    <div class="context-label">Balance As Of</div>
                    <div class="context-value">{format_date(balance_as_of_date)}</div>
                </div>
                <div class="context-card">
                    <div class="context-label">Planned Workbook</div>
                    <div class="context-value">{planned_source_name}</div>
                </div>
                <div class="context-card">
                    <div class="context-label">Outflow Categories</div>
                    <div class="context-value">{outflow_categories}</div>
                </div>
                <div class="context-card">
                    <div class="context-label">Holiday Region</div>
                    <div class="context-value">{holiday_region}</div>
                </div>
                <div class="context-card">
                    <div class="context-label">Actuals Status</div>
                    <div class="context-value">{actual_status}</div>
                </div>
                <div class="context-card">
                    <div class="context-label">TDS Rate Applied</div>
                    <div class="context-value">{tds_display}</div>
                </div>
                <div class="context-card">
                    <div class="context-label">Cash Floor</div>
                    <div class="context-value">{floor_display}</div>
                </div>
            </div>
            <div style="margin-top:10px; font-size:0.78rem; color:#9ca3af;">
                Workbook loaded: {load_ts_display}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def get_sample_sheet_names() -> list[str]:
    if not DEFAULT_SOURCE.exists():
        return []
    workbook = load_workbook(DEFAULT_SOURCE, read_only=True, data_only=False)
    return workbook.sheetnames


def render_upload_prompt() -> None:
    sample_sheets = get_sample_sheet_names()
    sheet_list = "".join(f"<li>{s}</li>" for s in sample_sheets) if sample_sheets else "<li>Receivables sheet</li><li>Separate outflow sheets by category</li>"
    st.title(APP_TITLE)
    st.caption("Upload today's workbook in the approved format to start the liquidity model.")
    st.markdown(
        f"""
        <div class="app-banner">
            <div class="app-banner-title">Planned workbook required</div>
            <div class="app-banner-subtitle">The app runs only on the workbook uploaded for the day.</div>
            <div style="margin-top:14px; font-size:0.94rem;">
                <strong>Expected format reference</strong>
                <ul style="margin-top:8px;">{sheet_list}</ul>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
#  UPLOAD HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

def handle_planned_upload_change() -> None:
    uploaded = st.session_state.get("planned_uploader_widget")
    st.session_state["planned_workbook_file"] = uploaded
    st.session_state["planned_load_time"] = datetime.now().strftime("%d-%m-%Y %H:%M:%S") if uploaded else None
    st.session_state.pop("selected_forecast_month", None)
    # Clear any in-app edits so fresh data seeds from the new file
    st.session_state.pop("_edit_source_hash", None)
    st.session_state.pop("edited_receivables", None)
    st.session_state.pop("edited_outflows", None)


def handle_actual_upload_change() -> None:
    uploaded = st.session_state.get("actual_uploader_widget")
    st.session_state["actual_workbook_file"] = uploaded


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")
    inject_styles()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title(APP_TITLE)
        st.caption("CFO-level liquidity control room")
        st.markdown("---")
        st.markdown("**Navigation**")
        selected_view = st.radio(
            "Select Section",
            ["Executive Summary", "Group Cashflow", "Collection Engine", "3-Month Cash Flow View", "Actual vs Budget", "✏️ Edit Data"],
            label_visibility="collapsed",
        )
        st.markdown("---")

        # ── Data Sources ──────────────────────────────────────────────────────
        st.markdown("**Data Sources**")
        st.file_uploader(
            "Planned workbook (Excel)", type=["xlsx", "xls"],
            key="planned_uploader_widget", on_change=handle_planned_upload_change,
        )
        st.file_uploader(
            "Actual workbook (Excel)", type=["xlsx", "xls"],
            key="actual_uploader_widget", on_change=handle_actual_upload_change,
        )
        st.markdown("---")

    # ── Load planned data ────────────────────────────────────────────────────
    planned_upload = st.session_state.get("planned_workbook_file", st.session_state.get("planned_uploader_widget"))
    if planned_upload is None:
        render_upload_prompt()
        return
    planned_source = planned_upload

    try:
        plan_data = load_data(planned_source, actual_mode=False)
    except Exception as exc:
        st.error(f"Unable to load planned workbook: {exc}")
        return

    # ── Seed / reconcile in-app edit state ──────────────────────────────────
    # On first load (or when file changes) copy raw Excel data into session
    # state.  Subsequent reruns use the session-state copy so in-app edits
    # survive without re-uploading.
    _current_hash = file_md5(planned_upload)
    if st.session_state.get("_edit_source_hash") != _current_hash:
        st.session_state["_edit_source_hash"] = _current_hash
        st.session_state["edited_receivables"] = plan_data["receivables"].copy()
        st.session_state["edited_outflows"] = plan_data["outflows"].copy()

    # Keep original Excel data available (used for reset in Edit Data section)
    _excel_receivables = plan_data["receivables"].copy()
    _excel_outflows = plan_data["outflows"].copy()

    # Override plan_data with session-state (possibly user-edited) versions
    plan_data["receivables"] = st.session_state["edited_receivables"].copy()
    plan_data["outflows"] = st.session_state["edited_outflows"].copy()
    # Keep outflow_sheet_order in sync with edited outflows
    if not plan_data["outflows"].empty:
        _edited_sheets = plan_data["outflows"]["sheet_name"].dropna().unique().tolist()
        plan_data["outflow_sheet_order"] = order_outflow_sheets(_edited_sheets)

    available_months = plan_data["available_months"]
    if not available_months:
        st.error("No valid forecast dates found in the planned workbook.")
        return

    # ── Sidebar: Forecast Settings ─────────────────────────────────────────
    with st.sidebar:
        if planned_upload:
            load_time = st.session_state.get("planned_load_time", "—")
            st.success(f"Planned: {planned_upload.name}")
            st.caption(f"Loaded: {load_time}")

        actual_upload = st.session_state.get("actual_workbook_file", st.session_state.get("actual_uploader_widget"))
        if actual_upload:
            st.success(f"Actual: {actual_upload.name}")
        else:
            st.caption("Upload an actual workbook to enable variance tracking.")

        st.markdown("**Forecast Settings**")
        selected_month = st.selectbox(
            "Forecast Month",
            options=available_months,
            format_func=month_label,
            index=default_month_index(available_months),
            key="selected_forecast_month",
        )
        current_bank_balance = st.number_input(
            "Current Bank Balance (₹)", min_value=0.0, value=5_000_000.0,
            step=100_000.0, format="%.0f",
        )
        holiday_region = st.selectbox(
            "Holiday Region", options=list(STATE_BANK_HOLIDAYS.keys()), index=0,
            help="Used for week-level bank-holiday highlighting.",
        )

        month_start = build_month_start(selected_month)
        month_end = build_month_end(selected_month)
        balance_date_key = "bank_balance_as_of_date"
        stored = st.session_state.get(balance_date_key)
        if stored is None:
            st.session_state[balance_date_key] = default_balance_as_of_date(selected_month).date()
        else:
            ts = pd.Timestamp(stored).normalize()
            if ts < month_start or ts > month_end:
                st.session_state[balance_date_key] = default_balance_as_of_date(selected_month).date()

        balance_as_of_input = st.date_input(
            "Bank Balance As Of Date",
            min_value=month_start.date(), max_value=month_end.date(),
            key=balance_date_key, format="DD-MM-YYYY",
            help="Forecast starts from this date.",
        )
        balance_as_of_date = pd.Timestamp(balance_as_of_input).normalize()
        st.markdown("---")

        # ── Model Assumptions ─────────────────────────────────────────────────
        st.markdown("**Model Assumptions**")

        # Minimum cash floor
        min_cash_threshold = st.number_input(
            "Minimum Cash Floor (₹)",
            min_value=0.0, value=DEFAULT_MIN_CASH_THRESHOLD,
            step=100_000.0, format="%.0f",
            help="Alert triggers if projected closing cash falls below this level.",
        )

        # TDS rate
        tds_pct = st.slider(
            "TDS Rate on Collections (%)", min_value=0.0, max_value=20.0,
            value=DEFAULT_TDS_RATE * 100, step=0.5,
            help="Percentage deducted at source on gross collection amounts. Net cash = Gross × (1 − TDS%).",
        )
        tds_rate = tds_pct / 100.0

        # Per-bucket collection probabilities
        st.markdown("**Collection Probabilities**")
        st.caption("Base = expected case. Worst = stress case. Adjust per your counterparty mix.")
        with st.expander("0–30 day bucket", expanded=False):
            b0_base = st.slider("Base %", 0, 100, int(DEFAULT_BASE_COLLECTION_PROBABILITIES["0-30"] * 100), key="prob_base_030") / 100
            b0_worst = st.slider("Worst %", 0, 100, int(DEFAULT_WORST_COLLECTION_PROBABILITIES["0-30"] * 100), key="prob_worst_030") / 100
        with st.expander("30–60 day bucket", expanded=False):
            b1_base = st.slider("Base %", 0, 100, int(DEFAULT_BASE_COLLECTION_PROBABILITIES["30-60"] * 100), key="prob_base_3060") / 100
            b1_worst = st.slider("Worst %", 0, 100, int(DEFAULT_WORST_COLLECTION_PROBABILITIES["30-60"] * 100), key="prob_worst_3060") / 100
        with st.expander("60+ day bucket", expanded=False):
            b2_base = st.slider("Base %", 0, 100, int(DEFAULT_BASE_COLLECTION_PROBABILITIES["60+"] * 100), key="prob_base_60p") / 100
            b2_worst = st.slider("Worst %", 0, 100, int(DEFAULT_WORST_COLLECTION_PROBABILITIES["60+"] * 100), key="prob_worst_60p") / 100

        user_base_probs = {"0-30": b0_base, "30-60": b1_base, "60+": b2_base}
        user_worst_probs = {"0-30": b0_worst, "30-60": b1_worst, "60+": b2_worst}
        st.markdown("---")

        # ── OD / Credit Line ──────────────────────────────────────────────────
        st.markdown("**OD / Credit Line**")
        od_sanctioned_limit = st.number_input(
            "OD Sanctioned Limit (₹)", min_value=0.0,
            value=DEFAULT_OD_LIMIT, step=500_000.0, format="%.0f",
            help="Total sanctioned overdraft / cash credit limit. Set to 0 if no OD facility.",
        )
        od_rate_pa = st.number_input(
            "OD Interest Rate (% p.a.)", min_value=0.0, max_value=50.0,
            value=DEFAULT_OD_RATE_PA, step=0.25, format="%.2f",
            help="Annual interest rate on OD drawn. Used to estimate monthly interest cost.",
        )
        st.markdown("---")

        # ── Model Context ─────────────────────────────────────────────────────
        st.markdown("**Model Context**")
        st.caption(f"Outflow sheets: {len(plan_data['outflow_sheet_order'])}")
        st.caption(f"Receivable lines: {len(plan_data['receivables'])}")
        if planned_upload:
            st.caption(f"File: {planned_upload.name}")
        load_time = st.session_state.get("planned_load_time", "—")
        st.caption(f"Loaded at: {load_time}")

    # ── Load actual data ────────────────────────────────────────────────────
    actual_upload = st.session_state.get("actual_workbook_file", st.session_state.get("actual_uploader_widget"))
    actual_data = None
    if actual_upload is not None:
        try:
            actual_data = load_data(actual_upload, actual_mode=True)
        except Exception as exc:
            st.error(f"Unable to load actual workbook: {exc}")
            return

    # ── Core computation ────────────────────────────────────────────────────
    weekly_results = compute_weekly_cashflow(
        data=plan_data,
        selected_month=selected_month,
        opening_balance=current_bank_balance,
        holiday_region=holiday_region,
        balance_as_of_date=balance_as_of_date,
        actual_data=actual_data,
        base_probs=user_base_probs,
        worst_probs=user_worst_probs,
        tds_rate=tds_rate,
        od_sanctioned_limit=od_sanctioned_limit,
        od_interest_rate_pa=od_rate_pa,
    )

    next_30_end = balance_as_of_date + pd.Timedelta(days=29)
    next_30_day_outflows = (
        filter_between(plan_data["outflows"], "date", balance_as_of_date, next_30_end)["amount"].sum()
        if not plan_data["outflows"].empty
        else 0.0
    )
    scenario_table, overdue_dependency, sixty_plus_dependency = calculate_collection_engine(
        plan_data["receivables"],
        balance_as_of_date,
        current_bank_balance,
        next_30_day_outflows,
        base_probs=user_base_probs,
        worst_probs=user_worst_probs,
        tds_rate=tds_rate,
    )

    planned_source_name = planned_upload.name
    actual_file_name = actual_upload.name if actual_upload else None
    load_ts = st.session_state.get("planned_load_time")

    # ── Page header ─────────────────────────────────────────────────────────
    st.title(APP_TITLE)
    st.caption("Structured liquidity management for weekly control, receivables sensitivity, and month-end cash visibility.")
    render_app_banner(
        selected_month=selected_month,
        balance_as_of_date=balance_as_of_date,
        planned_source_name=planned_source_name,
        outflow_categories=len(plan_data["outflow_sheet_order"]),
        holiday_region=holiday_region,
        actual_file_name=actual_file_name,
        load_timestamp=load_ts,
        tds_rate=tds_rate,
        min_cash_threshold=min_cash_threshold,
    )

    # ══ SECTION 1: Executive Summary ════════════════════════════════════════
    if selected_view == "Executive Summary":
        render_section_header(
            "Section 1: Executive Summary",
            "Month view of opening cash, net inflows, outflows, month-end balance, and liquidity runway.",
        )
        render_kpis(
            current_bank_balance,
            weekly_results,
            scenario_table,
            next_30_day_outflows,
            overdue_dependency,
            sixty_plus_dependency,
            min_cash_threshold=min_cash_threshold,
            od_sanctioned_limit=od_sanctioned_limit,
        )
        st.markdown("---")
        render_cfo_commentary("executive_summary")

    # ══ SECTION 2: Group Cashflow ════════════════════════════════════════════
    if selected_view == "Group Cashflow":
        render_section_header(
            "Section 2: Group Cashflow",
            "Weekly breakdown for the selected month with monthly forecast and actual columns.",
        )
        notes_col, dl_col, raw_col = st.columns([3, 1, 1])
        with notes_col:
            week_notes = " | ".join(w["label"] for w in weekly_results["week_meta"])
            st.caption(week_notes)
            st.caption(f"Forecast starts from {format_date(balance_as_of_date)}. TDS {tds_pct:.1f}% applied to collections.")
        with dl_col:
            export_bytes = build_weekly_breakdown_export(weekly_results, selected_month)
            st.download_button(
                "Download Weekly Breakdown (Excel)", data=export_bytes,
                file_name=f"weekly_cashflow_{selected_month.strftime('%Y_%m')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with raw_col:
            monthly_raw_bytes = build_month_raw_export(weekly_results, selected_month)
            st.download_button(
                "Download Monthly Raw Data (Excel)", data=monthly_raw_bytes,
                file_name=f"monthly_raw_data_{selected_month.strftime('%Y_%m')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        render_week_holiday_summary(weekly_results["week_meta"])
        financial_columns = weekly_results["week_columns"] + ["Fcst (Selected Month)", "Actual (if available)"]
        render_financial_table(weekly_results["weekly_table"], numeric_columns=financial_columns)
        render_weekly_drilldown(weekly_results["week_meta"])

    # ══ SECTION 3: Collection Engine ═════════════════════════════════════════
    if selected_view == "Collection Engine":
        render_section_header(
            "Section 3: Collection Engine",
            f"Receivables view using configured probabilities (Base {b0_base:.0%} / Worst {b0_worst:.0%} for 0-30 days). TDS {tds_pct:.1f}% applied.",
        )
        render_collection_engine(scenario_table, overdue_dependency, sixty_plus_dependency)

    # ══ SECTION 4: 3-Month Cash Flow View ═══════════════════════════════════
    if selected_view == "3-Month Cash Flow View":
        render_section_header(
            "Section 4: 3-Month Cash Flow View",
            f"Monthly aggregation for forward liquidity monitoring. Cash floor: {format_currency(min_cash_threshold) if min_cash_threshold > 0 else 'Not set'}.",
        )
        render_six_month_view(
            weekly_results["monthly_forecast"],
            weekly_results["monthly_inflow_detail"],
            weekly_results["monthly_outflow_detail"],
            min_cash_threshold=min_cash_threshold,
        )

    # ══ SECTION 5: Actual vs Budget ══════════════════════════════════════════
    if selected_view == "Actual vs Budget":
        render_section_header(
            "Section 5: Actual vs Budget",
            "Load actuals in the same workbook structure to compare against plan and isolate cash-impacting variances.",
        )
        render_actual_vs_budget(
            weekly_results,
            actual_data,
            actual_file_name,
            base_probs=user_base_probs,
            worst_probs=user_worst_probs,
            tds_rate=tds_rate,
        )


    # ══ SECTION 6: Edit Data ═════════════════════════════════════════════════
    if selected_view == "✏️ Edit Data":
        render_section_header(
            "Edit Data — Live Override",
            "Edit receivable lines or outflow entries directly. Every change is reflected immediately "
            "in all other sections without re-uploading Excel.",
        )

        _edit_badge = (
            "🟡 Working on edited data — edits active"
            if (
                not st.session_state["edited_receivables"].equals(_excel_receivables)
                or not st.session_state["edited_outflows"].equals(_excel_outflows)
            )
            else "🟢 Data matches the uploaded Excel"
        )
        st.caption(_edit_badge)

        tab_recv, tab_out = st.tabs(["📥 Receivables", "📤 Outflows"])

        # ── Tab 1: Receivables ────────────────────────────────────────────────
        with tab_recv:
            st.markdown(
                f"**{len(st.session_state['edited_receivables'])} receivable lines.** "
                "Edit amounts, dates, aging buckets or billing types. "
                "Add rows with ➕ or delete with the checkbox. "
                "Click **Apply** to commit changes."
            )
            recv_display_cols = [
                "counterparty", "invoice_date", "tentative_collection_date",
                "due_date", "amount", "ageing", "aging_bucket", "billing_bucket",
            ]
            recv_edit_df = (
                st.session_state["edited_receivables"]
                .reindex(columns=recv_display_cols)
                .copy()
            )
            # Ensure date columns are proper Python date for the editor
            for _dc in ["invoice_date", "tentative_collection_date", "due_date"]:
                recv_edit_df[_dc] = pd.to_datetime(recv_edit_df[_dc], errors="coerce").dt.date

            edited_recv_df = st.data_editor(
                recv_edit_df,
                key="recv_data_editor",
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "counterparty": st.column_config.TextColumn(
                        "Client / Party", help="Customer or counterparty name"
                    ),
                    "invoice_date": st.column_config.DateColumn(
                        "Invoice Date", format="DD/MM/YYYY"
                    ),
                    "tentative_collection_date": st.column_config.DateColumn(
                        "Expected Collection Date",
                        format="DD/MM/YYYY",
                        help="Drive when this receivable lands in the weekly forecast",
                    ),
                    "due_date": st.column_config.DateColumn(
                        "Due Date", format="DD/MM/YYYY"
                    ),
                    "amount": st.column_config.NumberColumn(
                        "Amount (₹)", format="%.0f", min_value=0
                    ),
                    "ageing": st.column_config.NumberColumn(
                        "Ageing (days)", format="%.0f", min_value=0
                    ),
                    "aging_bucket": st.column_config.SelectboxColumn(
                        "Aging Bucket", options=["0-30", "30-60", "60+"]
                    ),
                    "billing_bucket": st.column_config.SelectboxColumn(
                        "Billing Type", options=["Advance", "Arrear"]
                    ),
                },
            )

            _rc1, _rc2, _rc3 = st.columns([2, 1, 1])
            with _rc1:
                if st.button("✅ Apply Receivables Changes", use_container_width=True, type="primary"):
                    # Rebuild full receivables df from the editor output
                    _new_recv = edited_recv_df.copy()
                    for _dc in ["invoice_date", "tentative_collection_date", "due_date"]:
                        _new_recv[_dc] = pd.to_datetime(_new_recv[_dc], errors="coerce").dt.normalize()
                    _new_recv["amount"] = pd.to_numeric(_new_recv["amount"], errors="coerce").fillna(0.0)
                    _new_recv["ageing"] = pd.to_numeric(_new_recv["ageing"], errors="coerce").fillna(0.0)
                    _new_recv["counterparty"] = _new_recv["counterparty"].fillna("Receivable").astype(str)
                    _new_recv["aging_bucket"] = _new_recv["aging_bucket"].fillna(
                        _new_recv["ageing"].apply(classify_aging_bucket)
                    )
                    _new_recv["billing_bucket"] = _new_recv["billing_bucket"].fillna("Arrear")
                    # Rebuild description column to stay in sync
                    _new_recv["description"] = _new_recv["counterparty"]
                    # Drop zero-amount rows
                    _new_recv = _new_recv[_new_recv["amount"] != 0].reset_index(drop=True)
                    # Restore any columns that existed in original but are not editable here
                    _orig_cols = [c for c in st.session_state["edited_receivables"].columns
                                  if c not in recv_display_cols]
                    for _oc in _orig_cols:
                        _new_recv[_oc] = None
                    st.session_state["edited_receivables"] = _new_recv
                    st.success(f"✅ Receivables updated — {len(_new_recv)} lines active. All sections now reflect your changes.")
                    st.rerun()
            with _rc2:
                if st.button("🔄 Reset to Excel", use_container_width=True):
                    st.session_state["edited_receivables"] = _excel_receivables.copy()
                    st.success("Receivables reset to original Excel data.")
                    st.rerun()
            with _rc3:
                # Quick stats
                _total_recv = st.session_state["edited_receivables"]["amount"].sum()
                st.metric("Total Receivables (₹)", format_currency(_total_recv))

        # ── Tab 2: Outflows ───────────────────────────────────────────────────
        with tab_out:
            st.markdown(
                f"**{len(st.session_state['edited_outflows'])} outflow entries.** "
                "Edit amounts, dates, categories or add new rows. "
                "Click **Apply** to commit changes."
            )
            out_display_cols = ["sheet_name", "date", "vendor_name", "description", "amount"]
            out_edit_df = (
                st.session_state["edited_outflows"]
                .reindex(columns=out_display_cols)
                .copy()
            )
            out_edit_df["date"] = pd.to_datetime(out_edit_df["date"], errors="coerce").dt.date

            # Build list of known categories for the selectbox
            _known_cats = sorted(set(
                plan_data["outflow_sheet_order"]
                + st.session_state["edited_outflows"]["sheet_name"].dropna().unique().tolist()
            ))

            edited_out_df = st.data_editor(
                out_edit_df,
                key="out_data_editor",
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "sheet_name": st.column_config.SelectboxColumn(
                        "Category",
                        options=_known_cats,
                        help="Outflow category / sheet (e.g. Employee Salaries, Vendor Payment)",
                    ),
                    "date": st.column_config.DateColumn(
                        "Payment Date", format="DD/MM/YYYY"
                    ),
                    "vendor_name": st.column_config.TextColumn(
                        "Vendor / Party", help="Payee name"
                    ),
                    "description": st.column_config.TextColumn(
                        "Description", help="Nature of payment or invoice reference"
                    ),
                    "amount": st.column_config.NumberColumn(
                        "Amount (₹)", format="%.0f", min_value=0
                    ),
                },
            )

            _oc1, _oc2, _oc3 = st.columns([2, 1, 1])
            with _oc1:
                if st.button("✅ Apply Outflow Changes", use_container_width=True, type="primary"):
                    _new_out = edited_out_df.copy()
                    _new_out["date"] = pd.to_datetime(_new_out["date"], errors="coerce").dt.normalize()
                    _new_out["amount"] = pd.to_numeric(_new_out["amount"], errors="coerce").fillna(0.0)
                    _new_out["vendor_name"] = _new_out["vendor_name"].fillna("").astype(str)
                    _new_out["description"] = _new_out["description"].fillna("").astype(str)
                    _new_out["sheet_name"] = _new_out["sheet_name"].fillna("Other")
                    # Drop rows without date or amount
                    _new_out = _new_out[
                        _new_out["date"].notna() & (_new_out["amount"] != 0)
                    ].reset_index(drop=True)
                    st.session_state["edited_outflows"] = _new_out
                    st.success(f"✅ Outflows updated — {len(_new_out)} entries active. All sections now reflect your changes.")
                    st.rerun()
            with _oc2:
                if st.button("🔄 Reset to Excel ", use_container_width=True):
                    st.session_state["edited_outflows"] = _excel_outflows.copy()
                    st.success("Outflows reset to original Excel data.")
                    st.rerun()
            with _oc3:
                _total_out = st.session_state["edited_outflows"]["amount"].sum()
                st.metric("Total Outflows (₹)", format_currency(_total_out))

        st.markdown("---")
        st.markdown("#### Live Data Summary (post-edit)")
        _sum_c1, _sum_c2, _sum_c3 = st.columns(3)
        with _sum_c1:
            st.metric(
                "Receivable lines",
                len(st.session_state["edited_receivables"]),
                delta=len(st.session_state["edited_receivables"]) - len(_excel_receivables) or None,
            )
        with _sum_c2:
            st.metric(
                "Outflow entries",
                len(st.session_state["edited_outflows"]),
                delta=len(st.session_state["edited_outflows"]) - len(_excel_outflows) or None,
            )
        with _sum_c3:
            _net_chg = (
                st.session_state["edited_receivables"]["amount"].sum()
                - st.session_state["edited_outflows"]["amount"].sum()
            )
            st.metric("Net Position (₹)", format_currency(_net_chg))


if __name__ == "__main__":
    main()