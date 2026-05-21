"""
parser.py – Financial MIS parser.

Extracts CM and YTD P&L data from an MIS Excel workbook that contains:
  • List of Ledgers  – ledger master with classification & business vertical
  • TB               – current-month trial balance
  • TB YTD           – year-to-date trial balance
  • Stock            – stock movement (split into CM & YTD sections by marker row)
"""

import io
import re
import datetime
import pandas as pd
import numpy as np
from core.validation import validate_sheets_exist, validate_trial_balance, FinancialValidationError


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_header_row(df: pd.DataFrame, keyword: str) -> int:
    """Return the 0-based index of the first row that contains *keyword*."""
    for i in range(min(20, len(df))):
        if any(keyword.lower() in str(v).lower() for v in df.iloc[i]):
            return i
    return 0


def _read_with_header(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    """Promote the first row containing *keyword* to column headers."""
    hdr = _find_header_row(df, keyword)
    out = df.iloc[hdr:].copy()
    out.columns = out.iloc[0].astype(str).str.strip()
    out = out.iloc[1:].reset_index(drop=True)
    return out


def _get_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Return the first column name that matches any candidate (case-insensitive)."""
    for col in df.columns:
        col_str = str(col).strip()
        for cand in candidates:
            if cand.lower() in col_str.lower():
                return col
    return None


def _to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors='coerce').fillna(0)


# ---------------------------------------------------------------------------
# Stock sheet parser
# ---------------------------------------------------------------------------

def _parse_stock(raw: pd.DataFrame, report_month: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Split the Stock sheet into YTD and CM DataFrames.

    The sheet layout is:
        [header row] [YTD rows] [Total] [blank] [marker row with month/YTD label]
        [header row] [CM rows]  [Total]

    We detect the second header block by looking for a row whose first non-null
    cell contains the report-month string (e.g. "Mar'26") or "April to".
    """
    # Build a searchable string for the marker (e.g. "Mar'26")
    try:
        dt = datetime.datetime.strptime(report_month, "%Y-%m")
        short_month = dt.strftime("%b")          # "Mar"
        short_year  = dt.strftime("%y")          # "26"
        month_marker = f"{short_month}'{short_year}"   # "Mar'26"
    except ValueError:
        month_marker = ""

    # Find the row that separates the YTD summary (top) from the CM detail (bottom).
    # The CM section is always preceded by a row containing the short month label
    # (e.g. "Mar'26"). We deliberately do NOT match "April to" because that label
    # appears inside the YTD header area, not between the two sections.
    split_idx = None
    for i, row in raw.iterrows():
        row_str = " ".join(str(v) for v in row if pd.notna(v))
        if month_marker and month_marker.lower() in row_str.lower():
            split_idx = i
            break

    if split_idx is None:
        # Fallback: treat all as CM
        cm_raw  = raw
        ytd_raw = raw
    else:
        ytd_raw = raw.iloc[:split_idx]
        cm_raw  = raw.iloc[split_idx:]

    def _extract_section(section: pd.DataFrame) -> pd.DataFrame:
        hdr = _find_header_row(section.reset_index(drop=True), "Business Vert")
        sec = section.reset_index(drop=True).iloc[hdr:]
        sec.columns = sec.iloc[0].astype(str).str.strip()
        sec = sec.iloc[1:].reset_index(drop=True)

        # Normalise column names
        vert_col = _get_col(sec, ["business vert", "vertical", "verticle"])
        open_col = _get_col(sec, ["opening"])
        inwd_col = _get_col(sec, ["inward", "purchase"])
        outw_col = _get_col(sec, ["outward", "sales"])
        clos_col = _get_col(sec, ["closing"])

        rename = {}
        if vert_col: rename[vert_col] = "Business Vertical"
        if open_col: rename[open_col] = "Opening Stock"
        if inwd_col: rename[inwd_col] = "Purchases"
        if clos_col: rename[clos_col] = "Closing Stock"
        sec.rename(columns=rename, inplace=True)

        if "Business Vertical" not in sec.columns:
            return pd.DataFrame()

        sec = sec[sec["Business Vertical"].notna()].copy()
        sec = sec[~sec["Business Vertical"].astype(str).str.lower().isin(["total", "nan", ""])]
        sec["Business Vertical"] = sec["Business Vertical"].astype(str).str.strip().str.title()

        for col in ["Opening Stock", "Purchases", "Closing Stock"]:
            if col in sec.columns:
                sec[col] = _to_num(sec[col])

        keep = [c for c in ["Business Vertical", "Opening Stock", "Purchases", "Closing Stock"] if c in sec.columns]
        return sec[keep].groupby("Business Vertical", as_index=False).sum()

    return _extract_section(cm_raw), _extract_section(ytd_raw)


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_financial_excel(
    file_bytes: bytes,
    company_name: str = "Pristine Worldwide Private Limited",
    report_month: str = "2025-06",
) -> dict:
    try:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = excel_file.sheet_names

        required = ["List of Ledgers", "TB", "TB YTD", "Stock"]
        mapped   = validate_sheets_exist(sheet_names, required)

        # ------------------------------------------------------------------ #
        # 1.  Read raw sheets                                                 #
        # ------------------------------------------------------------------ #
        ledgers_raw  = pd.read_excel(excel_file, sheet_name=mapped["List of Ledgers"], header=None)
        tb_cm_raw    = pd.read_excel(excel_file, sheet_name=mapped["TB"],     header=None)
        tb_ytd_raw   = pd.read_excel(excel_file, sheet_name=mapped["TB YTD"], header=None)
        stock_raw    = pd.read_excel(excel_file, sheet_name=mapped["Stock"],  header=None)

        # ------------------------------------------------------------------ #
        # 2.  Parse Ledger master                                             #
        # ------------------------------------------------------------------ #
        ledgers = _read_with_header(ledgers_raw, "Name of Ledger")
        ledger_name_col = _get_col(ledgers, ["name of ledger", "ledger name"])
        vertical_col    = _get_col(ledgers, ["business vertical", "vertical"])
        group_col       = _get_col(ledgers, ["group"])
        head_col        = _get_col(ledgers, ["head"])

        if not ledger_name_col or not vertical_col:
            raise FinancialValidationError(
                title="Invalid Ledgers Format",
                message="'List of Ledgers' must have 'Name of Ledger' and 'Business Vertical' columns.",
            )

        rename = {ledger_name_col: "Ledger Name", vertical_col: "Business Vertical"}
        if group_col: rename[group_col] = "Group"
        if head_col:  rename[head_col]  = "Head"
        ledgers.rename(columns=rename, inplace=True)

        if "Group" not in ledgers.columns: ledgers["Group"] = ""
        if "Head"  not in ledgers.columns: ledgers["Head"]  = ""
        ledgers["Classification"] = (
            ledgers["Group"].astype(str) + " " + ledgers["Head"].astype(str)
        )
        ledgers = ledgers.dropna(subset=["Ledger Name"])
        ledgers = ledgers[ledgers["Ledger Name"].astype(str).str.strip() != ""]
        ledgers["Business Vertical"] = (
            ledgers["Business Vertical"].fillna("Unallocated")
            .astype(str).str.strip().str.title()
        )

        # ------------------------------------------------------------------ #
        # 3.  Parse CM Trial Balance                                          #
        # ------------------------------------------------------------------ #
        tb_cm = _read_with_header(tb_cm_raw, "Particulars")

        # Normalise: keep only the "Debit / Credit" columns (not "Debit Bal" etc.)
        # The TB sheet has: Opening Bal | Debit Bal | Credit Bal | Closing Bal | gap |
        #                   Opening     | Debit     | Credit     | Closing
        # We want the second set (Opening, Debit, Credit, Closing).
        part_col = _get_col(tb_cm, ["particulars", "ledger"])
        if not part_col:
            tb_cm.columns.values[0] = "Ledger Name"
        else:
            tb_cm.rename(columns={part_col: "Ledger Name"}, inplace=True)

        # Find the exact "Debit" and "Credit" columns (not "Debit Bal")
        debit_col  = None
        credit_col = None
        open_col   = None
        close_col  = None
        for col in tb_cm.columns:
            cs = str(col).strip()
            if cs == "Debit":   debit_col  = col
            if cs == "Credit":  credit_col = col
            if cs == "Opening": open_col   = col
            if cs == "Closing": close_col  = col

        # Fallback: "Debit Bal" → "Debit", "Credit Bal" → "Credit" etc.
        if not debit_col:
            debit_col  = _get_col(tb_cm, ["debit"])
        if not credit_col:
            credit_col = _get_col(tb_cm, ["credit"])
        if not open_col:
            open_col   = _get_col(tb_cm, ["opening"])
        if not close_col:
            close_col  = _get_col(tb_cm, ["closing"])

        tb_cm_clean = tb_cm[["Ledger Name"]].copy()
        if open_col:   tb_cm_clean["Opening"] = _to_num(tb_cm[open_col])
        if debit_col:  tb_cm_clean["Debit"]   = _to_num(tb_cm[debit_col])
        if credit_col: tb_cm_clean["Credit"]  = _to_num(tb_cm[credit_col])
        if close_col:  tb_cm_clean["Closing"] = _to_num(tb_cm[close_col])

        validate_trial_balance(tb_cm_clean)

        # ------------------------------------------------------------------ #
        # 4.  Parse YTD Trial Balance                                         #
        # ------------------------------------------------------------------ #
        tb_ytd = _read_with_header(tb_ytd_raw, "Particulars")

        part_col_y = _get_col(tb_ytd, ["particulars", "ledger"])
        if not part_col_y:
            tb_ytd.columns.values[0] = "Ledger Name"
        else:
            tb_ytd.rename(columns={part_col_y: "Ledger Name"}, inplace=True)

        # YTD sheet has: Opening | Debit | Credit | Closing | gap |
        #                Opening YTD | Debit YTD | Credit YTD | Closing YTD
        # We want the YTD columns (the second set).
        ytd_debit  = _get_col(tb_ytd, ["debit ytd"])
        ytd_credit = _get_col(tb_ytd, ["credit ytd"])
        ytd_open   = _get_col(tb_ytd, ["opening ytd"])
        ytd_close  = _get_col(tb_ytd, ["closing ytd"])

        # Fallback to first Debit/Credit if YTD columns absent
        if not ytd_debit:  ytd_debit  = _get_col(tb_ytd, ["debit"])
        if not ytd_credit: ytd_credit = _get_col(tb_ytd, ["credit"])
        if not ytd_open:   ytd_open   = _get_col(tb_ytd, ["opening"])
        if not ytd_close:  ytd_close  = _get_col(tb_ytd, ["closing"])

        tb_ytd_clean = tb_ytd[["Ledger Name"]].copy()
        if ytd_open:   tb_ytd_clean["Opening YTD"]  = _to_num(tb_ytd[ytd_open])
        if ytd_debit:  tb_ytd_clean["Debit YTD"]    = _to_num(tb_ytd[ytd_debit])
        if ytd_credit: tb_ytd_clean["Credit YTD"]   = _to_num(tb_ytd[ytd_credit])
        if ytd_close:  tb_ytd_clean["Closing YTD"]  = _to_num(tb_ytd[ytd_close])

        # ------------------------------------------------------------------ #
        # 5.  Merge: CM TB ←→ YTD TB ←→ Ledger master                       #
        # ------------------------------------------------------------------ #
        merged = pd.merge(tb_cm_clean, tb_ytd_clean, on="Ledger Name", how="outer")

        ledger_cols = ["Ledger Name", "Business Vertical", "Classification"]
        merged = pd.merge(
            merged,
            ledgers[ledger_cols].drop_duplicates(subset=["Ledger Name"]),
            on="Ledger Name",
            how="left",
        )

        # Fill defaults
        for col in ["Debit", "Credit", "Opening", "Closing",
                    "Debit YTD", "Credit YTD", "Opening YTD", "Closing YTD"]:
            if col not in merged.columns:
                merged[col] = 0.0
            else:
                merged[col] = _to_num(merged[col])

        merged["Business Vertical"] = (
            merged["Business Vertical"].fillna("Unallocated")
            .astype(str).str.strip().str.title()
        )
        merged["Classification"] = merged["Classification"].fillna("").astype(str)

        # Net Balance convention:
        #   Sales accounts  → Credit > Debit → positive Net Balance = revenue
        #   Expense/Purchase → Debit > Credit → we negate at point-of-use
        merged["Net Balance"]     = merged["Credit"]     - merged["Debit"]
        merged["Net Balance YTD"] = merged["Credit YTD"] - merged["Debit YTD"]

        # ------------------------------------------------------------------ #
        # 6.  Vertical order (from Ledger master, first-seen)                 #
        # ------------------------------------------------------------------ #
        seen: set[str] = set()
        vertical_order: list[str] = []
        for v in ledgers["Business Vertical"]:
            if v not in seen:
                seen.add(v)
                vertical_order.append(v)

        verticals = [v for v in vertical_order if v in merged["Business Vertical"].values]

        # ------------------------------------------------------------------ #
        # 7.  Month display strings                                           #
        # ------------------------------------------------------------------ #
        try:
            dt = datetime.datetime.strptime(report_month, "%Y-%m")
            month_name         = dt.strftime("%B")
            year_name          = dt.strftime("%Y")
            report_month_disp  = f"{month_name} {year_name}"
            ytd_range          = f"(April to {month_name})"
        except ValueError:
            report_month_disp  = report_month
            ytd_range          = "(YTD)"

        # ------------------------------------------------------------------ #
        # 8.  Parse Stock sheet                                               #
        # ------------------------------------------------------------------ #
        stock_cm, stock_ytd = _parse_stock(stock_raw, report_month)

        # ------------------------------------------------------------------ #
        # 9.  Build report_data                                               #
        # ------------------------------------------------------------------ #
        report_data: dict = {
            "company_name":   company_name,
            "report_month":   report_month_disp,
            "ytd_range":      ytd_range,
            "vertical_order": vertical_order,
            "vertical_types": {},
            "verticals":      {},
            "summary": {
                "total_revenue":      0.0,
                "total_cogs":         0.0,
                "total_gross_profit": 0.0,
            },
        }

        # ------------------------------------------------------------------ #
        # 10. Per-vertical calculations                                       #
        # ------------------------------------------------------------------ #
        for vertical in verticals:
            vd = merged[merged["Business Vertical"] == vertical]

            def _clf(pattern: str, exclude: str = "") -> pd.DataFrame:
                mask = vd["Classification"].str.contains(pattern, case=False, na=False, regex=True)
                if exclude:
                    mask &= ~vd["Classification"].str.contains(exclude, case=False, na=False, regex=True)
                return vd[mask]

            # --- Sales (matches "1. Sales Accounts", "Sales Accounts", "Sales Account", etc.) ---
            sales_rows  = _clf(r"Sales\s*Account")
            revenue     = sales_rows["Net Balance"].sum()        # credit accounts → positive
            ytd_revenue = sales_rows["Net Balance YTD"].sum()

            # --- Purchases (matches "5. Purchase Accounts", "Purchase Accounts", etc.) ---
            purch_rows       = _clf(r"Purchase\s*Account")
            tb_purchases     = -purch_rows["Net Balance"].sum()       # debit-heavy → flip → positive
            ytd_tb_purchases = -purch_rows["Net Balance YTD"].sum()

            # --- Stock (CM) ---
            v_stk_cm = stock_cm[stock_cm["Business Vertical"] == vertical] if not stock_cm.empty else pd.DataFrame()
            opening_stock  = float(v_stk_cm["Opening Stock"].sum())  if "Opening Stock" in v_stk_cm.columns else 0.0
            closing_stock  = float(v_stk_cm["Closing Stock"].sum())  if "Closing Stock" in v_stk_cm.columns else 0.0
            stock_purch_cm = float(v_stk_cm["Purchases"].sum())      if "Purchases"     in v_stk_cm.columns else 0.0
            purchases      = tb_purchases if tb_purchases != 0 else stock_purch_cm
            cogs           = opening_stock + purchases - closing_stock

            # --- Stock (YTD) ---
            v_stk_ytd = stock_ytd[stock_ytd["Business Vertical"] == vertical] if not stock_ytd.empty else pd.DataFrame()
            ytd_opening = float(v_stk_ytd["Opening Stock"].sum()) if "Opening Stock" in v_stk_ytd.columns else 0.0
            ytd_closing = float(v_stk_ytd["Closing Stock"].sum()) if "Closing Stock" in v_stk_ytd.columns else 0.0
            ytd_stk_pur = float(v_stk_ytd["Purchases"].sum())     if "Purchases"     in v_stk_ytd.columns else 0.0
            ytd_purchases = ytd_tb_purchases if ytd_tb_purchases != 0 else ytd_stk_pur
            ytd_cogs      = ytd_opening + ytd_purchases - ytd_closing

            # --- Direct Expenses (matches "3. Direct Expense", "Direct Expense", etc.) ---
            dir_rows = _clf(r"Direct\s*Expense", exclude=r"Indirect")
            direct_expenses     = -dir_rows["Net Balance"].sum()
            ytd_direct_expenses = -dir_rows["Net Balance YTD"].sum()
            direct_breakdown: dict[str, float] = {}
            ytd_direct_breakdown: dict[str, float] = {}
            for _, row in dir_rows.iterrows():
                lname = str(row["Ledger Name"])
                cm_val  = -float(row["Net Balance"])
                ytd_val = -float(row["Net Balance YTD"])
                if cm_val  != 0: direct_breakdown[lname]     = cm_val
                if ytd_val != 0: ytd_direct_breakdown[lname] = ytd_val

            gross_profit     = revenue     - cogs     - direct_expenses
            ytd_gross_profit = ytd_revenue - ytd_cogs - ytd_direct_expenses

            # --- Indirect Income (matches "2. Indirect Income", "Indirect Income", etc.) ---
            inc_rows = _clf(r"Indirect\s*Income")
            indirect_income     = inc_rows["Net Balance"].sum()
            ytd_indirect_income = inc_rows["Net Balance YTD"].sum()
            income_breakdown: dict[str, float] = {}
            ytd_income_breakdown: dict[str, float] = {}
            for _, row in inc_rows.iterrows():
                lname = str(row["Ledger Name"])
                cm_val  = float(row["Net Balance"])
                ytd_val = float(row["Net Balance YTD"])
                if cm_val  != 0: income_breakdown[lname]     = cm_val
                if ytd_val != 0: ytd_income_breakdown[lname] = ytd_val

            # --- Indirect Expenses (matches "6. Indirect Expense", "Indirect Expense", etc.) ---
            # Exclude Indirect Income rows so they are not double-counted.
            ind_rows = _clf(r"Indirect\s*Expense", exclude=r"Indirect\s*Income|Direct")
            indirect_expenses     = -ind_rows["Net Balance"].sum()
            ytd_indirect_expenses = -ind_rows["Net Balance YTD"].sum()
            indirect_breakdown: dict[str, float] = {}
            ytd_indirect_breakdown: dict[str, float] = {}
            for _, row in ind_rows.iterrows():
                lname = str(row["Ledger Name"])
                cm_val  = -float(row["Net Balance"])
                ytd_val = -float(row["Net Balance YTD"])
                if cm_val  != 0: indirect_breakdown[lname]     = cm_val
                if ytd_val != 0: ytd_indirect_breakdown[lname] = ytd_val

            net_profit     = gross_profit     + indirect_income     - indirect_expenses
            ytd_net_profit = ytd_gross_profit + ytd_indirect_income - ytd_indirect_expenses

            # --- Debtors / Creditors ---
            def _balance_dict(rows: pd.DataFrame) -> dict:
                def _s(col: str) -> float:
                    return float(_to_num(rows[col]).sum()) if col in rows.columns else 0.0
                return {
                    "opening":     _s("Opening"),
                    "debit":       _s("Debit"),
                    "credit":      _s("Credit"),
                    "closing":     _s("Closing"),
                    "opening_ytd": _s("Opening YTD"),
                    "debit_ytd":   _s("Debit YTD"),
                    "credit_ytd":  _s("Credit YTD"),
                    "closing_ytd": _s("Closing YTD"),
                }

            # Keyword-based matching for Debtors/Creditors works for any Tally export
            # ("Sundry Debtors", "Debtors", "Sundry Debtor" etc.)
            dr_rows = _clf(r"Debtor")
            cr_rows = _clf(r"Creditor")

            # --- Vertical type ---
            has_revenue = any([revenue, cogs, direct_expenses, ytd_revenue, ytd_cogs, ytd_direct_expenses])
            v_lower = vertical.lower()
            if "share" in v_lower and "trading" in v_lower:
                v_type = "share_trading"
            elif not has_revenue:
                v_type = "cost_center"
            else:
                v_type = "revenue"

            report_data["vertical_types"][vertical] = v_type
            report_data["verticals"][vertical] = {
                "revenue":          float(revenue),
                "cogs":             float(cogs),
                "direct_expenses":  float(direct_expenses),
                "direct_breakdown": direct_breakdown,
                "gross_profit":     float(gross_profit),
                "indirect_income":  float(indirect_income),
                "income_breakdown": income_breakdown,
                "indirect_expenses":  float(indirect_expenses),
                "indirect_breakdown": indirect_breakdown,
                "net_profit":         float(net_profit),
                "details": {
                    "opening_stock": float(opening_stock),
                    "purchases":     float(purchases),
                    "closing_stock": float(closing_stock),
                },
                # YTD
                "ytd_revenue":             float(ytd_revenue),
                "ytd_cogs":                float(ytd_cogs),
                "ytd_direct_expenses":     float(ytd_direct_expenses),
                "ytd_direct_breakdown":    ytd_direct_breakdown,
                "ytd_gross_profit":        float(ytd_gross_profit),
                "ytd_indirect_income":     float(ytd_indirect_income),
                "ytd_income_breakdown":    ytd_income_breakdown,
                "ytd_indirect_expenses":   float(ytd_indirect_expenses),
                "ytd_indirect_breakdown":  ytd_indirect_breakdown,
                "ytd_net_profit":          float(ytd_net_profit),
                "debtors":   _balance_dict(dr_rows),
                "creditors": _balance_dict(cr_rows),
            }

            report_data["summary"]["total_revenue"]      += revenue
            report_data["summary"]["total_cogs"]         += cogs
            report_data["summary"]["total_gross_profit"] += gross_profit

        return report_data

    except FinancialValidationError:
        raise
    except Exception as exc:
        raise FinancialValidationError(
            title="Parsing Error",
            message="An error occurred while parsing the Excel file.",
            details=[str(exc)],
        )
