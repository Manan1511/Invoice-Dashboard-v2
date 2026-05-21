import io
import pandas as pd
import numpy as np
import datetime
from core.validation import validate_sheets_exist, validate_trial_balance, FinancialValidationError

def get_header_row(df, keyword):
    for i in range(min(15, len(df))):
        row_values = df.iloc[i].astype(str).str.lower()
        if any(keyword.lower() in str(val) for val in row_values):
            df.columns = df.iloc[i]
            df = df.iloc[i+1:].reset_index(drop=True)
            df.columns = df.columns.astype(str).str.strip()
            return df
    df.columns = df.columns.astype(str).str.strip()
    return df

def get_col(df, possible_names):
    for col in df.columns:
        for name in possible_names:
            if name.lower() in col.lower():
                return col
    return None

def parse_financial_excel(file_bytes: bytes, company_name: str = "Pristine Worldwide Private Limited", report_month: str = "2025-06") -> dict:
    try:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = excel_file.sheet_names
        
        required_sheets = ["List of Ledgers", "TB", "Stock"]
        mapped_sheets = validate_sheets_exist(sheet_names, required_sheets)
        
        ledgers_df = pd.read_excel(excel_file, sheet_name=mapped_sheets["List of Ledgers"], header=None)
        tb_df = pd.read_excel(excel_file, sheet_name=mapped_sheets["TB"], header=None)
        stock_df = pd.read_excel(excel_file, sheet_name=mapped_sheets["Stock"], header=None)
        
        ledgers_df = get_header_row(ledgers_df, "Name of Ledger")
        tb_df = get_header_row(tb_df, "Debit")
        stock_df = get_header_row(stock_df, "Opening")
        
        # Rename columns to standardized names for internal use
        tb_debit_col = get_col(tb_df, ["debit"])
        tb_credit_col = get_col(tb_df, ["credit"])
        
        if tb_debit_col and tb_debit_col != 'Debit':
            # Handle duplicate column names by dropping the existing 'Debit' if it exists
            if 'Debit' in tb_df.columns:
                tb_df = tb_df.loc[:, ~tb_df.columns.duplicated()]
                tb_df = tb_df.drop(columns=['Debit'])
            tb_df.rename(columns={tb_debit_col: 'Debit'}, inplace=True)
            
        if tb_credit_col and tb_credit_col != 'Credit':
            if 'Credit' in tb_df.columns:
                tb_df = tb_df.loc[:, ~tb_df.columns.duplicated()]
                tb_df = tb_df.drop(columns=['Credit'])
            tb_df.rename(columns={tb_credit_col: 'Credit'}, inplace=True)
            
        validate_trial_balance(tb_df)
        
        # Determine Ledger columns
        ledger_name_col = get_col(ledgers_df, ["name of ledger", "ledger name"])
        vertical_col = get_col(ledgers_df, ["business vertical", "vertical"])
        group_col = get_col(ledgers_df, ["group"])
        head_col = get_col(ledgers_df, ["head"])
        
        if not ledger_name_col or not vertical_col:
            raise FinancialValidationError(
                title="Invalid Ledgers Format",
                message="The 'List of Ledgers' sheet must contain Ledger Name and Business Vertical columns.",
            )
            
        rename_dict = {
            ledger_name_col: 'Ledger Name',
            vertical_col: 'Business Vertical',
        }
        if group_col: rename_dict[group_col] = 'Group'
        if head_col: rename_dict[head_col] = 'Head'
            
        ledgers_df.rename(columns=rename_dict, inplace=True)
        
        # Drop rows where Ledger Name is empty or whitespace
        ledgers_df = ledgers_df.dropna(subset=['Ledger Name'])
        ledgers_df = ledgers_df[ledgers_df['Ledger Name'].astype(str).str.strip() != '']
        
        # Combine Group and Head for classification
        if 'Group' not in ledgers_df.columns: ledgers_df['Group'] = ''
        if 'Head' not in ledgers_df.columns: ledgers_df['Head'] = ''
        ledgers_df['Classification'] = ledgers_df['Group'].astype(str) + " " + ledgers_df['Head'].astype(str)
        
        # Gracefully handle missing Business Verticals by assigning them to 'Unallocated'
        ledgers_df['Business Vertical'] = ledgers_df['Business Vertical'].fillna('Unallocated')
        
        # Normalize essential column names if present
        if "Name of Ledger" in ledgers_df.columns:
            ledgers_df.rename(columns={"Name of Ledger": "Ledger Name"}, inplace=True)
            
        tb_ledger_col = get_col(tb_df, ["particulars", "ledger"])
        if tb_ledger_col:
            tb_df.rename(columns={tb_ledger_col: "Ledger Name"}, inplace=True)
        else:
            tb_df.rename(columns={tb_df.columns[0]: "Ledger Name"}, inplace=True)
            
        # Create Classification group
        ledgers_df['Classification'] = ledgers_df['Group'].astype(str) + ' ' + ledgers_df['Head'].astype(str)
        
        # Merge TB with Ledgers to map classifications, verticals, AND grab YTD columns!
        ytd_cols = []
        if 'Opening YTD' in ledgers_df.columns: ytd_cols.extend(['Opening YTD', 'Debit YTD', 'Credit YTD', 'Closing YTD'])
        merge_cols = ['Ledger Name', 'Business Vertical', 'Classification'] + ytd_cols
        
        merged_tb = pd.merge(tb_df, ledgers_df[[c for c in merge_cols if c in ledgers_df.columns]], on='Ledger Name', how='left')
        
        merged_tb['Net Balance'] = pd.to_numeric(merged_tb['Credit'], errors='coerce').fillna(0) - pd.to_numeric(merged_tb['Debit'], errors='coerce').fillna(0)
        if 'Credit YTD' in merged_tb.columns:
            merged_tb['YTD Net Balance'] = pd.to_numeric(merged_tb['Credit YTD'], errors='coerce').fillna(0) - pd.to_numeric(merged_tb['Debit YTD'], errors='coerce').fillna(0)
        else:
            merged_tb['YTD Net Balance'] = 0
            
        # Normalize verticals (capitalize properly, e.g. 'common' -> 'Common')
        ledgers_df['Business Vertical'] = ledgers_df['Business Vertical'].astype(str).str.strip().str.title()
        merged_tb['Business Vertical'] = merged_tb['Business Vertical'].astype(str).str.strip().str.title()
        
        # Preserve the exact order verticals first appear in the List of Ledgers sheet
        seen = set()
        vertical_order = []
        for v in ledgers_df['Business Vertical']:
            v = str(v).strip()
            if pd.notna(v) and v != '' and v != 'Nan' and v not in seen:
                seen.add(v)
                vertical_order.append(v)
        verticals = vertical_order
        
        # Helper to safely parse month string to get exact names
        try:
            date_obj = datetime.datetime.strptime(report_month, "%Y-%m")
            month_name = date_obj.strftime("%B")
            year_name = date_obj.strftime("%Y")
            report_month_display = f"{month_name} {year_name}"
            # e.g., 'May 2025' -> YTD range 'April to May'
            ytd_range = f"(April to {month_name})"
        except:
            report_month_display = report_month
            ytd_range = "(YTD)"
            
        report_data = {
            "company_name": company_name,
            "report_month": report_month_display,
            "ytd_range": ytd_range,
            "vertical_order": vertical_order,  # natural order from List of Ledgers
            "vertical_types": {},              # filled in per-vertical loop below
            "verticals": {},
            "summary": {
                "total_revenue": 0,
                "total_cogs": 0,
                "total_gross_profit": 0
            },
            "raw_summary": {
                "total_rows": len(merged_tb)
            }
        }
        
        # Stock headers mapping
        stock_vert_col = get_col(stock_df, ["vertical", "verticle"])
        if stock_vert_col: stock_df.rename(columns={stock_vert_col: 'Business Vertical'}, inplace=True)
        
        stock_opening_col = get_col(stock_df, ["opening"])
        stock_closing_col = get_col(stock_df, ["closing"])
        stock_purchases_col = get_col(stock_df, ["inward", "purchase"])
        
        if stock_opening_col: stock_df.rename(columns={stock_opening_col: 'Opening Stock'}, inplace=True)
        if stock_closing_col: stock_df.rename(columns={stock_closing_col: 'Closing Stock'}, inplace=True)
        if stock_purchases_col: stock_df.rename(columns={stock_purchases_col: 'Purchases'}, inplace=True)
        
        # The stock sheet has YTD totals (top) and Current Month totals (bottom). 
        if 'Business Vertical' in stock_df.columns:
            stock_df['Business Vertical'] = stock_df['Business Vertical'].astype(str).str.strip().str.title()
            stock_df_cm = stock_df.drop_duplicates(subset=['Business Vertical'], keep='last')
            stock_df_ytd = stock_df.drop_duplicates(subset=['Business Vertical'], keep='first')
        else:
            stock_df_cm = pd.DataFrame()
            stock_df_ytd = pd.DataFrame()
            
        for vertical in verticals:
            v_data = merged_tb[merged_tb['Business Vertical'] == vertical]
            
            # Helper to get sum safely
            def sum_col(df, col_name):
                if col_name in df.columns:
                    return pd.to_numeric(df[col_name], errors='coerce').fillna(0).sum()
                return 0
                
            # Current Month calculations
            sales_data = v_data[v_data['Classification'].astype(str).str.contains('Sales', case=False, na=False)]
            revenue = sales_data['Net Balance'].sum()
            
            purchases_data = v_data[v_data['Classification'].astype(str).str.contains('Purchase', case=False, na=False)]
            tb_purchases = -purchases_data['Net Balance'].sum()
            
            v_stock_cm = stock_df_cm[stock_df_cm['Business Vertical'] == vertical] if not stock_df_cm.empty else pd.DataFrame()
            opening_stock = pd.to_numeric(v_stock_cm['Opening Stock'], errors='coerce').sum() if 'Opening Stock' in v_stock_cm.columns else 0
            closing_stock = pd.to_numeric(v_stock_cm['Closing Stock'], errors='coerce').sum() if 'Closing Stock' in v_stock_cm.columns else 0
            stock_purchases = pd.to_numeric(v_stock_cm['Purchases'], errors='coerce').sum() if 'Purchases' in v_stock_cm.columns else 0
            
            purchases = tb_purchases if tb_purchases != 0 else stock_purchases
            cogs = opening_stock + purchases - closing_stock
            # Direct Expenses
            direct_data = v_data[v_data['Classification'].astype(str).str.contains('Direct', case=False, na=False)]
            direct_data = direct_data[~direct_data['Classification'].astype(str).str.contains('Indirect', case=False, na=False)]
            direct_expenses = -direct_data['Net Balance'].sum()
            direct_breakdown = {}
            for _, row in direct_data.iterrows():
                val = -float(row['Net Balance'])
                if val != 0: direct_breakdown[row['Ledger Name']] = val
            
            gross_profit = revenue - cogs - direct_expenses
            
            indirect_inc_data = v_data[v_data['Classification'].astype(str).str.contains('Indirect Income', case=False, na=False)]
            indirect_income = indirect_inc_data['Net Balance'].sum()
            
            indirect_data = v_data[v_data['Classification'].astype(str).str.contains('Indirect|Expense|Factory|Office|Common', case=False, na=False)]
            indirect_data = indirect_data[~indirect_data['Classification'].astype(str).str.contains('Indirect Income|Direct', case=False, na=False)]
            
            indirect_expenses = -indirect_data['Net Balance'].sum()
            
            indirect_breakdown = {}
            for _, row in indirect_data.iterrows():
                val = -float(row['Net Balance'])
                if val != 0: indirect_breakdown[row['Ledger Name']] = val
                    
            income_breakdown = {}
            for _, row in indirect_inc_data.iterrows():
                val = float(row['Net Balance'])
                if val != 0: income_breakdown[row['Ledger Name']] = val
            
            net_profit = gross_profit + indirect_income - indirect_expenses
            
            # YTD Calculations
            ytd_revenue = sales_data['YTD Net Balance'].sum()
            ytd_tb_purchases = -purchases_data['YTD Net Balance'].sum()
            
            v_stock_ytd = stock_df_ytd[stock_df_ytd['Business Vertical'] == vertical] if not stock_df_ytd.empty else pd.DataFrame()
            ytd_opening_stock = pd.to_numeric(v_stock_ytd['Opening Stock'], errors='coerce').sum() if 'Opening Stock' in v_stock_ytd.columns else 0
            ytd_closing_stock = pd.to_numeric(v_stock_ytd['Closing Stock'], errors='coerce').sum() if 'Closing Stock' in v_stock_ytd.columns else 0
            ytd_stock_purchases = pd.to_numeric(v_stock_ytd['Purchases'], errors='coerce').sum() if 'Purchases' in v_stock_ytd.columns else 0
            
            ytd_purchases = ytd_tb_purchases if ytd_tb_purchases != 0 else ytd_stock_purchases
            ytd_cogs = ytd_opening_stock + ytd_purchases - ytd_closing_stock
            
            ytd_direct_expenses = -direct_data['YTD Net Balance'].sum()
            ytd_direct_breakdown = {}
            for _, row in direct_data.iterrows():
                val = -float(row['YTD Net Balance'])
                if val != 0: ytd_direct_breakdown[row['Ledger Name']] = val
                
            ytd_gross_profit = ytd_revenue - ytd_cogs - ytd_direct_expenses
            
            ytd_indirect_income = indirect_inc_data['YTD Net Balance'].sum()
            ytd_indirect_expenses = -indirect_data['YTD Net Balance'].sum()
            
            ytd_indirect_breakdown = {}
            for _, row in indirect_data.iterrows():
                val = -float(row['YTD Net Balance'])
                if val != 0: ytd_indirect_breakdown[row['Ledger Name']] = val
                
            ytd_net_profit = ytd_gross_profit + ytd_indirect_income - ytd_indirect_expenses
            
            debtors_data = v_data[v_data['Classification'].astype(str).str.contains('Sundry Debtor', case=False, na=False)]
            creditors_data = v_data[v_data['Classification'].astype(str).str.contains('Sundry Creditor', case=False, na=False)]
            
            report_data["verticals"][vertical] = {
                "revenue": float(revenue),
                "cogs": float(cogs),
                "direct_expenses": float(direct_expenses),
                "direct_breakdown": direct_breakdown,
                "gross_profit": float(gross_profit),
                "indirect_income": float(indirect_income),
                "indirect_expenses": float(indirect_expenses),
                "net_profit": float(net_profit),
                "indirect_breakdown": indirect_breakdown,
                "income_breakdown": income_breakdown,
                "details": {
                    "opening_stock": float(opening_stock),
                    "purchases": float(purchases),
                    "closing_stock": float(closing_stock)
                },
                "ytd_revenue": float(ytd_revenue),
                "ytd_cogs": float(ytd_cogs),
                "ytd_direct_expenses": float(ytd_direct_expenses),
                "ytd_direct_breakdown": ytd_direct_breakdown,
                "ytd_gross_profit": float(ytd_gross_profit),
                "ytd_indirect_income": float(ytd_indirect_income),
                "ytd_indirect_expenses": float(ytd_indirect_expenses),
                "ytd_net_profit": float(ytd_net_profit),
                "ytd_indirect_breakdown": ytd_indirect_breakdown,
                "debtors": {
                    "opening": float(sum_col(debtors_data, 'Opening')),
                    "debit": float(sum_col(debtors_data, 'Debit')),
                    "credit": float(sum_col(debtors_data, 'Credit')),
                    "closing": float(sum_col(debtors_data, 'Closing')),
                    "opening_ytd": float(sum_col(debtors_data, 'Opening YTD')),
                    "debit_ytd": float(sum_col(debtors_data, 'Debit YTD')),
                    "credit_ytd": float(sum_col(debtors_data, 'Credit YTD')),
                    "closing_ytd": float(sum_col(debtors_data, 'Closing YTD'))
                },
                "creditors": {
                    "opening": float(sum_col(creditors_data, 'Opening')),
                    "debit": float(sum_col(creditors_data, 'Debit')),
                    "credit": float(sum_col(creditors_data, 'Credit')),
                    "closing": float(sum_col(creditors_data, 'Closing')),
                    "opening_ytd": float(sum_col(creditors_data, 'Opening YTD')),
                    "debit_ytd": float(sum_col(creditors_data, 'Debit YTD')),
                    "credit_ytd": float(sum_col(creditors_data, 'Credit YTD')),
                    "closing_ytd": float(sum_col(creditors_data, 'Closing YTD'))
                }
            }
            
            # Classify the vertical so the generator needs zero hardcoded names.
            # A vertical whose name contains 'share' + 'trading' is share_trading.
            # A vertical with zero revenue AND zero cogs in BOTH CM and YTD is a cost_center.
            # Everything else is a revenue-generating vertical.
            v_lower = vertical.lower()
            has_revenue = (revenue != 0 or cogs != 0 or direct_expenses != 0 or
                           ytd_revenue != 0 or ytd_cogs != 0 or ytd_direct_expenses != 0)
            if 'share' in v_lower and 'trading' in v_lower:
                v_type = 'share_trading'
            elif not has_revenue:
                v_type = 'cost_center'
            else:
                v_type = 'revenue'
            report_data["vertical_types"][vertical] = v_type

            report_data["summary"]["total_revenue"] += revenue
            report_data["summary"]["total_cogs"] += cogs
            report_data["summary"]["total_gross_profit"] += gross_profit
            
        return report_data
        
    except FinancialValidationError:
        raise
    except Exception as e:
        raise FinancialValidationError(
            title="Parsing Error",
            message="An error occurred while parsing the Excel file.",
            details=[str(e)]
        )
