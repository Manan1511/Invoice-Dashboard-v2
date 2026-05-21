import pandas as pd
import tempfile
import uuid
import os

def generate_excel_report(report_data: dict) -> str:
    """
    Generates an Excel file with P&L and Drs_Crs format sheets based on report_data.
    Returns the absolute file path to the generated temp file.
    """
    temp_dir = os.path.join(os.getcwd(), "temp_reports")
    os.makedirs(temp_dir, exist_ok=True)
    
    filename = f"report_{uuid.uuid4().hex}.xlsx"
    filepath = os.path.join(temp_dir, filename)
    
    verticals = list(report_data["verticals"].keys())
    
    # ----------------------------------------------------
    # 1. P&L Sheet Data
    # ----------------------------------------------------
    # We want a matrix: Rows = Line items, Columns = Verticals + Total
    pl_rows = [
        "Sales",
        "Less: COGS",
        "Gross margin",
        "Gross margin %",
        "Indirect income",
        "Depreciation & amortization",
        "Net income",
        "Net allocable income",
        "Net profit"
    ]
    
    pl_data = { "Particulars": pl_rows }
    
    total_sales = 0
    total_cogs = 0
    total_gross = 0
    total_indirect = 0
    total_net = 0
    
    for vertical in verticals:
        v_data = report_data["verticals"][vertical]
        sales = v_data["revenue"]
        cogs = v_data["cogs"]
        gross = v_data["gross_profit"]
        indirect = v_data["indirect_expenses"]
        net = v_data["net_profit"]
        
        gross_margin_pct = (gross / sales) if sales != 0 else 0
        
        pl_data[vertical] = [
            sales,
            cogs,
            gross,
            gross_margin_pct,
            0, # Indirect income not separated yet
            0, # Dep & amort
            gross, # Net income before indirect
            0, # Net allocable
            net
        ]
        
        total_sales += sales
        total_cogs += cogs
        total_gross += gross
        total_indirect += indirect
        total_net += net
        
    total_gross_pct = (total_gross / total_sales) if total_sales != 0 else 0
    pl_data["Total"] = [
        total_sales,
        total_cogs,
        total_gross,
        total_gross_pct,
        0,
        0,
        total_gross,
        0,
        total_net
    ]
    
    pl_df = pd.DataFrame(pl_data)
    
    # ----------------------------------------------------
    # 2. Drs_Crs Sheet Data
    # ----------------------------------------------------
    drcrs_rows = []
    
    # Debtors section
    drcrs_rows.append({"Category": "Sundry Debtor", "Opening": "Opening", "Debit": "Debit", "Credit": "Credit", "Closing": "Closing"})
    tot_dr_open = tot_dr_debit = tot_dr_credit = tot_dr_closing = 0
    for vertical in verticals:
        v_data = report_data["verticals"][vertical]["debtors"]
        drcrs_rows.append({
            "Category": vertical,
            "Opening": v_data["opening"],
            "Debit": v_data["debit"],
            "Credit": v_data["credit"],
            "Closing": v_data["closing"]
        })
        tot_dr_open += v_data["opening"]
        tot_dr_debit += v_data["debit"]
        tot_dr_credit += v_data["credit"]
        tot_dr_closing += v_data["closing"]
        
    drcrs_rows.append({"Category": "Grand Total", "Opening": tot_dr_open, "Debit": tot_dr_debit, "Credit": tot_dr_credit, "Closing": tot_dr_closing})
    drcrs_rows.append({"Category": "", "Opening": "", "Debit": "", "Credit": "", "Closing": ""}) # blank row
    
    # Creditors section
    drcrs_rows.append({"Category": "Sundry Creditor", "Opening": "Opening", "Debit": "Debit", "Credit": "Credit", "Closing": "Closing"})
    tot_cr_open = tot_cr_debit = tot_cr_credit = tot_cr_closing = 0
    for vertical in verticals:
        v_data = report_data["verticals"][vertical]["creditors"]
        drcrs_rows.append({
            "Category": vertical,
            "Opening": v_data["opening"],
            "Debit": v_data["debit"],
            "Credit": v_data["credit"],
            "Closing": v_data["closing"]
        })
        tot_cr_open += v_data["opening"]
        tot_cr_debit += v_data["debit"]
        tot_cr_credit += v_data["credit"]
        tot_cr_closing += v_data["closing"]
        
    drcrs_rows.append({"Category": "Grand Total", "Opening": tot_cr_open, "Debit": tot_cr_debit, "Credit": tot_cr_credit, "Closing": tot_cr_closing})
    
    drcrs_df = pd.DataFrame(drcrs_rows)
    drcrs_df.rename(columns={"Category": "", "Opening": "", "Debit": "", "Credit": "", "Closing": ""}, inplace=True)
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        pl_df.to_excel(writer, sheet_name='P&L', index=False, startrow=4)
        
        # Write headers for P&L manually
        worksheet_pl = writer.sheets['P&L']
        worksheet_pl.cell(row=1, column=1, value="Pristine Worldwide Private Limited")
        worksheet_pl.cell(row=3, column=1, value="P&L Analysis")
        worksheet_pl.cell(row=4, column=1, value="For the period")
        
        drcrs_df.to_excel(writer, sheet_name='Drs_Crs', index=False, startrow=3)
        worksheet_drcr = writer.sheets['Drs_Crs']
        worksheet_drcr.cell(row=1, column=1, value="Pristine Worldwide Private Limited")
        worksheet_drcr.cell(row=3, column=1, value="Summary of Sundry Debtors & Sundry Creditors")
        
    return filename
