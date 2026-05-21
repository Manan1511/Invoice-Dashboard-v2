import io
import pandas as pd
import numpy as np
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

def parse_financial_excel(file_bytes: bytes) -> dict:
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
        
        # Ensure TB has Ledger Name
        tb_ledger_col = get_col(tb_df, ["particular", "ledger", "name", "nan"])
        if tb_ledger_col:
            tb_df.rename(columns={tb_ledger_col: 'Ledger Name'}, inplace=True)
        else:
            tb_df.rename(columns={tb_df.columns[0]: 'Ledger Name'}, inplace=True)
            
        merged_tb = pd.merge(tb_df, ledgers_df[['Ledger Name', 'Business Vertical', 'Classification']], on='Ledger Name', how='left')
        
        merged_tb['Net Balance'] = pd.to_numeric(merged_tb['Credit'], errors='coerce').fillna(0) - pd.to_numeric(merged_tb['Debit'], errors='coerce').fillna(0)
        
        verticals = [v for v in ledgers_df['Business Vertical'].unique() if pd.notna(v) and str(v).strip() != '']
        
        report_data = {
            "verticals": {},
            "summary": {
                "total_revenue": 0,
                "total_cogs": 0,
                "total_gross_profit": 0
            }
        }
        
        # Normalize stock columns
        stock_vertical_col = get_col(stock_df, ["business verticle", "business vertical", "vertical"])
        stock_opening_col = get_col(stock_df, ["opening balance", "opening stock"])
        stock_closing_col = get_col(stock_df, ["closing balance", "closing stock"])
        stock_purchases_col = get_col(stock_df, ["inward", "purchase"])
        
        if stock_vertical_col: stock_df.rename(columns={stock_vertical_col: 'Business Vertical'}, inplace=True)
        if stock_opening_col: stock_df.rename(columns={stock_opening_col: 'Opening Stock'}, inplace=True)
        if stock_closing_col: stock_df.rename(columns={stock_closing_col: 'Closing Stock'}, inplace=True)
        if stock_purchases_col: stock_df.rename(columns={stock_purchases_col: 'Purchases'}, inplace=True)
        
        for vertical in verticals:
            v_data = merged_tb[merged_tb['Business Vertical'] == vertical]
            
            sales_data = v_data[v_data['Classification'].astype(str).str.contains('Sales|Revenue', case=False, na=False)]
            revenue = sales_data['Net Balance'].sum()
            
            purchases_data = v_data[v_data['Classification'].astype(str).str.contains('Purchase', case=False, na=False)]
            tb_purchases = -purchases_data['Net Balance'].sum()
            
            v_stock = stock_df[stock_df['Business Vertical'] == vertical] if 'Business Vertical' in stock_df.columns else pd.DataFrame()
            opening_stock = pd.to_numeric(v_stock['Opening Stock'], errors='coerce').sum() if 'Opening Stock' in v_stock.columns else 0
            closing_stock = pd.to_numeric(v_stock['Closing Stock'], errors='coerce').sum() if 'Closing Stock' in v_stock.columns else 0
            stock_purchases = pd.to_numeric(v_stock['Purchases'], errors='coerce').sum() if 'Purchases' in v_stock.columns else 0
            
            # Prefer purchases from TB, fallback to Stock Inward
            purchases = tb_purchases if tb_purchases != 0 else stock_purchases
            
            cogs = opening_stock + purchases - closing_stock
            gross_profit = revenue - cogs
            
            indirect_data = v_data[v_data['Classification'].astype(str).str.contains('Indirect|Expense|Factory|Office|Common', case=False, na=False)]
            indirect_expenses = -indirect_data['Net Balance'].sum()
            
            net_profit = gross_profit - indirect_expenses
            
            # Aggregate Debtors and Creditors for Drs_Crs sheet
            debtors_data = v_data[v_data['Classification'].astype(str).str.contains('Sundry Debtor', case=False, na=False)]
            creditors_data = v_data[v_data['Classification'].astype(str).str.contains('Sundry Creditor', case=False, na=False)]
            
            # Helper to get sum safely
            def sum_col(df, col_name):
                if col_name in df.columns:
                    return pd.to_numeric(df[col_name], errors='coerce').fillna(0).sum()
                return 0
                
            report_data["verticals"][vertical] = {
                "revenue": float(revenue),
                "cogs": float(cogs),
                "gross_profit": float(gross_profit),
                "indirect_expenses": float(indirect_expenses),
                "net_profit": float(net_profit),
                "details": {
                    "opening_stock": float(opening_stock),
                    "purchases": float(purchases),
                    "closing_stock": float(closing_stock)
                },
                "debtors": {
                    "opening": float(sum_col(debtors_data, 'Opening Bal')),
                    "debit": float(sum_col(debtors_data, 'Debit')),
                    "credit": float(sum_col(debtors_data, 'Credit')),
                    "closing": float(sum_col(debtors_data, 'Closing Bal'))
                },
                "creditors": {
                    "opening": float(sum_col(creditors_data, 'Opening Bal')),
                    "debit": float(sum_col(creditors_data, 'Debit')),
                    "credit": float(sum_col(creditors_data, 'Credit')),
                    "closing": float(sum_col(creditors_data, 'Closing Bal'))
                }
            }
            
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
