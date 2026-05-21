import pandas as pd
import tempfile
import uuid
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

def generate_excel_report(report_data: dict) -> str:
    temp_dir = os.path.join(os.getcwd(), "temp_reports")
    os.makedirs(temp_dir, exist_ok=True)
    
    filename = f"report_{uuid.uuid4().hex}.xlsx"
    filepath = os.path.join(temp_dir, filename)
    
    company_name = report_data.get("company_name", "Pristine Worldwide Private Limited")
    report_month = report_data.get("report_month", "May 2025")
    ytd_range = report_data.get("ytd_range", "(April to May)")
    
    # Use the natural order from the List of Ledgers sheet (emitted by the parser)
    # Fall back to dict key order if the parser is an older version
    vertical_order = report_data.get("vertical_order", list(report_data["verticals"].keys()))
    vertical_types = report_data.get("vertical_types", {})

    # If vertical_types is missing (old parser), derive it from data
    if not vertical_types:
        for v, d in report_data["verticals"].items():
            v_lower = v.lower()
            if 'share' in v_lower and 'trading' in v_lower:
                vertical_types[v] = 'share_trading'
            elif d['revenue'] == 0 and d['cogs'] == 0 and d.get('direct_expenses', 0) == 0:
                vertical_types[v] = 'cost_center'
            else:
                vertical_types[v] = 'revenue'

    # Respect the parser order; only include verticals that actually exist in the data
    ordered_verticals = [v for v in vertical_order if v in report_data["verticals"]]
    # Append any that the parser emitted but weren't in vertical_order (safety)
    for v in report_data["verticals"]:
        if v not in ordered_verticals:
            ordered_verticals.append(v)

    # Drop purely-zero Unallocated verticals
    if 'Unallocated' in ordered_verticals:
        d = report_data["verticals"]['Unallocated']
        if d['revenue'] == 0 and d['cogs'] == 0 and d['indirect_expenses'] == 0 and d['indirect_income'] == 0:
            ordered_verticals.remove('Unallocated')

    # Classify from the type map
    rev_verticals = [v for v in ordered_verticals if vertical_types.get(v) == 'revenue']
    cost_centers = [v for v in ordered_verticals if vertical_types.get(v) == 'cost_center']
    share_trading_verticals = [v for v in ordered_verticals if vertical_types.get(v) == 'share_trading']

    # Fallback: if nothing classified as revenue, treat all non-cost-center non-share-trading as revenue
    if not rev_verticals:
        rev_verticals = [v for v in ordered_verticals if v not in cost_centers and v not in share_trading_verticals]
            
    # Gather indirect and direct ledgers (CM + YTD)
    all_indirect_ledgers = set()
    all_direct_ledgers = set()
    for v_data in report_data["verticals"].values():
        all_indirect_ledgers.update(v_data.get("indirect_breakdown", {}).keys())
        all_indirect_ledgers.update(v_data.get("ytd_indirect_breakdown", {}).keys())
        all_direct_ledgers.update(v_data.get("direct_breakdown", {}).keys())
        all_direct_ledgers.update(v_data.get("ytd_direct_breakdown", {}).keys())
    all_indirect_ledgers = sorted(list(all_indirect_ledgers))
    all_direct_ledgers = sorted(list(all_direct_ledgers))
    
    if "Depreciation & amortization" in all_indirect_ledgers: all_indirect_ledgers.remove("Depreciation & amortization")
    if "Depreciation" in all_indirect_ledgers: all_indirect_ledgers.remove("Depreciation")
        
    pl_rows = ["Sales", "Less: COGS"]
    if all_direct_ledgers:
        pl_rows.append("3. Direct Expense")
        pl_rows.extend(all_direct_ledgers)
        
    pl_rows.extend([
        "Gross margin", "Gross margin %", "Gross margin % (previous month)", "",
        "Indirect income", "Depreciation & amortization", "Net income", "Net allocable income", "", "Net profit", "",
        "Indirect costs:"
    ])
    pl_rows.extend(all_indirect_ledgers)
    pl_rows.extend([
        "Total indirect costs", "", "Allocation of expenses:"
    ])
    pl_rows.extend(cost_centers)
    pl_rows.extend([
        "Total allocable expenses", "", "Profit/ (loss) before tax", "Net margin %", "Net margin % (previous month)", "", "",
        "Profit/(loss) as per tally", "Difference"
    ])
    
    pl_data = { "Particulars": pl_rows }
    
    def build_col(v_data, is_ytd=False):
        prefix = "ytd_" if is_ytd else ""
        sales = v_data.get(f"{prefix}revenue", 0)
        cogs = v_data.get(f"{prefix}cogs", 0)
        dir_exps = v_data.get(f"{prefix}direct_expenses", 0)
        gross = sales - cogs - dir_exps
        
        ind_inc = v_data.get(f"{prefix}indirect_income", 0)
        depr = v_data.get(f"{prefix}indirect_breakdown", {}).get("Depreciation & amortization", 0) + \
               v_data.get(f"{prefix}indirect_breakdown", {}).get("Depreciation", 0)
        
        net_inc = ind_inc
        net_alloc = 0
        net_prof = gross + net_inc - depr
        
        gross_margin_pct = (gross / sales) if sales != 0 else 0
        
        col_data = [
            sales if sales != 0 else "-", cogs if cogs != 0 else "-"
        ]
        
        if all_direct_ledgers:
            col_data.append(dir_exps if dir_exps != 0 else "-")
            for ledger in all_direct_ledgers:
                val = v_data.get(f"{prefix}direct_breakdown", {}).get(ledger, 0)
                col_data.append(val if val != 0 else 0)
                
        col_data.extend([
            gross if gross != 0 else "-",
            f"{gross_margin_pct*100:.2f}%" if sales != 0 else "0.00%", "0.00%", "",
            ind_inc if ind_inc != 0 else "-", depr if depr != 0 else "-", net_inc if net_inc != 0 else "-",
            net_alloc if net_alloc != 0 else "-", "", net_prof if net_prof != 0 else "-", "", ""
        ])
        
        tot_ind_costs = 0
        for ledger in all_indirect_ledgers:
            val = v_data.get(f"{prefix}indirect_breakdown", {}).get(ledger, 0)
            col_data.append(val if val != 0 else 0)
            tot_ind_costs += val
            
        col_data.append(tot_ind_costs if tot_ind_costs != 0 else "-")
        col_data.extend(["", ""]) # blank, allocation header
        
        tot_alloc = 0
        for cc in cost_centers:
            cc_costs = sum(report_data["verticals"][cc].get(f"{prefix}indirect_breakdown", {}).values())
            allocated_val = 0
            if v in rev_verticals:
                allocated_val = cc_costs / len(rev_verticals)
            elif v == cc:
                allocated_val = -cc_costs
            col_data.append(allocated_val if allocated_val != 0 else "-")
            tot_alloc += allocated_val
            
        col_data.append(tot_alloc if tot_alloc != 0 else "-")
        col_data.append("")
        
        profit_before_tax = net_prof - tot_ind_costs - tot_alloc
        col_data.append(profit_before_tax if profit_before_tax != 0 else "-")
        net_margin_pct = (profit_before_tax / sales) if sales != 0 else 0
        col_data.append(f"{net_margin_pct*100:.2f}%" if sales != 0 else "0.00%")
        col_data.append("0.00%") # previous month margin
        
        col_data.extend(["", "", profit_before_tax if profit_before_tax != 0 else "-", "-"])
        return col_data, sales, cogs, dir_exps, gross, ind_inc, depr, net_inc, net_prof, tot_ind_costs, tot_alloc, profit_before_tax
        
    def build_totals(all_sales, all_cogs, all_dir, all_gross, all_ind_inc, all_depr, all_net_inc, all_net_prof, all_tot_ind, all_profit, is_ytd, breakdown_totals, dir_breakdown_totals, alloc_totals):
        gross_pct = (all_gross/all_sales) if all_sales != 0 else 0
        col = [
            all_sales if all_sales!=0 else "-", all_cogs if all_cogs!=0 else "-"
        ]
        
        if all_direct_ledgers:
            col.append(all_dir if all_dir != 0 else "-")
            for l in all_direct_ledgers:
                col.append(dir_breakdown_totals[l] if dir_breakdown_totals[l]!=0 else 0)
                
        col.extend([
            all_gross if all_gross!=0 else "-",
            f"{gross_pct*100:.2f}%" if all_sales!=0 else "0.00%", "0.00%", "",
            all_ind_inc if all_ind_inc!=0 else "-", all_depr if all_depr!=0 else "-", all_net_inc if all_net_inc!=0 else "-",
            "-", "", all_net_prof if all_net_prof!=0 else "-", "", ""
        ])
        
        for l in all_indirect_ledgers: col.append(breakdown_totals[l] if breakdown_totals[l]!=0 else 0)
        col.append(all_tot_ind if all_tot_ind!=0 else "-")
        col.extend(["", ""])
        for cc in cost_centers: col.append("-")
        col.append("-")
        col.append("")
        col.append(all_profit if all_profit!=0 else "-")
        net_pct = (all_profit/all_sales) if all_sales!=0 else 0
        col.append(f"{net_pct*100:.2f}%" if all_sales!=0 else "0.00%")
        col.append("0.00%")
        col.extend(["", "", all_profit if all_profit!=0 else "-", "-"])
        return col

    cm_totals_no_st = {
        'sales':0, 'cogs':0, 'dir':0, 'gross':0, 'ind_inc':0, 'depr':0, 'net_inc':0, 'net_prof':0, 
        'tot_ind':0, 'profit':0, 'breakdown': {l:0 for l in all_indirect_ledgers}, 'dir_breakdown': {l:0 for l in all_direct_ledgers}, 'alloc': {cc:0 for cc in cost_centers}
    }
    cm_totals_inc_st = {
        'sales':0, 'cogs':0, 'dir':0, 'gross':0, 'ind_inc':0, 'depr':0, 'net_inc':0, 'net_prof':0, 
        'tot_ind':0, 'profit':0, 'breakdown': {l:0 for l in all_indirect_ledgers}, 'dir_breakdown': {l:0 for l in all_direct_ledgers}, 'alloc': {cc:0 for cc in cost_centers}
    }
    
    st_col_cm = None
    
    for v in ordered_verticals:
        if v in share_trading_verticals: continue
        v_data = report_data["verticals"][v]
        col, s, c, dx, g, ii, d, ni, np, ti, ta, pbt = build_col(v_data, False)
        pl_data[v] = col
        
        for t_dict in (cm_totals_no_st, cm_totals_inc_st):
            t_dict['sales']+=s; t_dict['cogs']+=c; t_dict['dir']+=dx; t_dict['gross']+=g; t_dict['ind_inc']+=ii; t_dict['depr']+=d
            t_dict['net_inc']+=ni; t_dict['net_prof']+=np; t_dict['tot_ind']+=ti; t_dict['profit']+=pbt
            for l in all_indirect_ledgers: t_dict['breakdown'][l] += v_data.get("indirect_breakdown", {}).get(l, 0)
            for l in all_direct_ledgers: t_dict['dir_breakdown'][l] += v_data.get("direct_breakdown", {}).get(l, 0)
            
    pl_data["Total (without share trading)"] = build_totals(
        cm_totals_no_st['sales'], cm_totals_no_st['cogs'], cm_totals_no_st['dir'], cm_totals_no_st['gross'], cm_totals_no_st['ind_inc'],
        cm_totals_no_st['depr'], cm_totals_no_st['net_inc'], cm_totals_no_st['net_prof'], cm_totals_no_st['tot_ind'],
        cm_totals_no_st['profit'], False, cm_totals_no_st['breakdown'], cm_totals_no_st['dir_breakdown'], cm_totals_no_st['alloc'])
        
    for st_v in share_trading_verticals:
        st_data = report_data["verticals"][st_v]
        st_col_cm, s, c, dx, g, ii, d, ni, np, ti, ta, pbt = build_col(st_data, False)
        pl_data[st_v] = st_col_cm
        t_dict = cm_totals_inc_st
        t_dict['sales']+=s; t_dict['cogs']+=c; t_dict['dir']+=dx; t_dict['gross']+=g; t_dict['ind_inc']+=ii; t_dict['depr']+=d
        t_dict['net_inc']+=ni; t_dict['net_prof']+=np; t_dict['tot_ind']+=ti; t_dict['profit']+=pbt
        for l in all_indirect_ledgers: t_dict['breakdown'][l] += st_data.get("indirect_breakdown", {}).get(l, 0)
        for l in all_direct_ledgers: t_dict['dir_breakdown'][l] += st_data.get("direct_breakdown", {}).get(l, 0)
        
    pl_data["Total (including share trading)"] = build_totals(
        cm_totals_inc_st['sales'], cm_totals_inc_st['cogs'], cm_totals_inc_st['dir'], cm_totals_inc_st['gross'], cm_totals_inc_st['ind_inc'],
        cm_totals_inc_st['depr'], cm_totals_inc_st['net_inc'], cm_totals_inc_st['net_prof'], cm_totals_inc_st['tot_ind'],
        cm_totals_inc_st['profit'], False, cm_totals_inc_st['breakdown'], cm_totals_inc_st['dir_breakdown'], cm_totals_inc_st['alloc'])
    
    pl_data[" "] = [""] * len(pl_rows)
    
    ytd_totals_no_st = {
        'sales':0, 'cogs':0, 'dir':0, 'gross':0, 'ind_inc':0, 'depr':0, 'net_inc':0, 'net_prof':0, 
        'tot_ind':0, 'profit':0, 'breakdown': {l:0 for l in all_indirect_ledgers}, 'dir_breakdown': {l:0 for l in all_direct_ledgers}, 'alloc': {cc:0 for cc in cost_centers}
    }
    ytd_totals_inc_st = {
        'sales':0, 'cogs':0, 'dir':0, 'gross':0, 'ind_inc':0, 'depr':0, 'net_inc':0, 'net_prof':0, 
        'tot_ind':0, 'profit':0, 'breakdown': {l:0 for l in all_indirect_ledgers}, 'dir_breakdown': {l:0 for l in all_direct_ledgers}, 'alloc': {cc:0 for cc in cost_centers}
    }
    
    st_col_ytd = None
    
    for v in ordered_verticals:
        if v in share_trading_verticals: continue
        v_data = report_data["verticals"][v]
        col, s, c, dx, g, ii, d, ni, np, ti, ta, pbt = build_col(v_data, True)
        pl_data[v+"_YTD"] = col
        
        for t_dict in (ytd_totals_no_st, ytd_totals_inc_st):
            t_dict['sales']+=s; t_dict['cogs']+=c; t_dict['dir']+=dx; t_dict['gross']+=g; t_dict['ind_inc']+=ii; t_dict['depr']+=d
            t_dict['net_inc']+=ni; t_dict['net_prof']+=np; t_dict['tot_ind']+=ti; t_dict['profit']+=pbt
            for l in all_indirect_ledgers: t_dict['breakdown'][l] += v_data.get("ytd_indirect_breakdown", {}).get(l, 0)
            for l in all_direct_ledgers: t_dict['dir_breakdown'][l] += v_data.get("ytd_direct_breakdown", {}).get(l, 0)
            
    pl_data["Total (without share trading)_YTD"] = build_totals(
        ytd_totals_no_st['sales'], ytd_totals_no_st['cogs'], ytd_totals_no_st['dir'], ytd_totals_no_st['gross'], ytd_totals_no_st['ind_inc'],
        ytd_totals_no_st['depr'], ytd_totals_no_st['net_inc'], ytd_totals_no_st['net_prof'], ytd_totals_no_st['tot_ind'],
        ytd_totals_no_st['profit'], True, ytd_totals_no_st['breakdown'], ytd_totals_no_st['dir_breakdown'], ytd_totals_no_st['alloc'])
        
    for st_v in share_trading_verticals:
        st_data = report_data["verticals"][st_v]
        st_col_ytd, s, c, dx, g, ii, d, ni, np, ti, ta, pbt = build_col(st_data, True)
        pl_data[st_v+"_YTD"] = st_col_ytd
        t_dict = ytd_totals_inc_st
        t_dict['sales']+=s; t_dict['cogs']+=c; t_dict['dir']+=dx; t_dict['gross']+=g; t_dict['ind_inc']+=ii; t_dict['depr']+=d
        t_dict['net_inc']+=ni; t_dict['net_prof']+=np; t_dict['tot_ind']+=ti; t_dict['profit']+=pbt
        for l in all_indirect_ledgers: t_dict['breakdown'][l] += st_data.get("ytd_indirect_breakdown", {}).get(l, 0)
        for l in all_direct_ledgers: t_dict['dir_breakdown'][l] += st_data.get("ytd_direct_breakdown", {}).get(l, 0)
        
    pl_data["Total (including share trading)_YTD"] = build_totals(
        ytd_totals_inc_st['sales'], ytd_totals_inc_st['cogs'], ytd_totals_inc_st['dir'], ytd_totals_inc_st['gross'], ytd_totals_inc_st['ind_inc'],
        ytd_totals_inc_st['depr'], ytd_totals_inc_st['net_inc'], ytd_totals_inc_st['net_prof'], ytd_totals_inc_st['tot_ind'],
        ytd_totals_inc_st['profit'], True, ytd_totals_inc_st['breakdown'], ytd_totals_inc_st['dir_breakdown'], ytd_totals_inc_st['alloc'])
    
    pl_df = pd.DataFrame(pl_data)
    
    drcrs_rows = [{"Cat1": "Sundry Debtor", "Cat2": "Opening", "Cat3": "Debit", "Cat4": "Credit", "Cat5": "Closing", "Cat6": "", "Cat7": "Opening", "Cat8": "Debit", "Cat9": "Credit", "Cat10": "Closing"}]
    tod, tdd, tcd, tcld, tody, tddy, tcdy, tcldy = 0,0,0,0,0,0,0,0
    for v in ordered_verticals:
        d = report_data["verticals"].get(v, {}).get("debtors", {"opening":0, "debit":0, "credit":0, "closing":0, "opening_ytd":0, "debit_ytd":0, "credit_ytd":0, "closing_ytd":0})
        if d["opening"]==0 and d["debit"]==0 and d["credit"]==0 and d["closing"]==0 and d["opening_ytd"]==0 and d["debit_ytd"]==0 and d["credit_ytd"]==0 and d["closing_ytd"]==0:
            continue
        drcrs_rows.append({"Cat1": v, "Cat2": d["opening"], "Cat3": d["debit"], "Cat4": d["credit"], "Cat5": d["closing"], "Cat6": "", "Cat7": d["opening_ytd"], "Cat8": d["debit_ytd"], "Cat9": d["credit_ytd"], "Cat10": d["closing_ytd"]})
        tod+=d["opening"]; tdd+=d["debit"]; tcd+=d["credit"]; tcld+=d["closing"]
        tody+=d["opening_ytd"]; tddy+=d["debit_ytd"]; tcdy+=d["credit_ytd"]; tcldy+=d["closing_ytd"]
        
    drcrs_rows.append({"Cat1": "Grand Total", "Cat2": tod, "Cat3": tdd, "Cat4": tcd, "Cat5": tcld, "Cat6": "", "Cat7": tody, "Cat8": tddy, "Cat9": tcdy, "Cat10": tcldy})
    drcrs_rows.append({"Cat1": "", "Cat2": "", "Cat3": "", "Cat4": "", "Cat5": "", "Cat6": "", "Cat7": "", "Cat8": "", "Cat9": "", "Cat10": ""})
    
    drcrs_rows.append({"Cat1": "Sundry Creditor", "Cat2": "Opening", "Cat3": "Debit", "Cat4": "Credit", "Cat5": "Closing", "Cat6": "", "Cat7": "Opening", "Cat8": "Debit", "Cat9": "Credit", "Cat10": "Closing"})
    toc, tdc, tcc, tclc, tocy, tdcy, tccy, tclcy = 0,0,0,0,0,0,0,0
    for v in ordered_verticals:
        d = report_data["verticals"].get(v, {}).get("creditors", {"opening":0, "debit":0, "credit":0, "closing":0, "opening_ytd":0, "debit_ytd":0, "credit_ytd":0, "closing_ytd":0})
        if d["opening"]==0 and d["debit"]==0 and d["credit"]==0 and d["closing"]==0 and d["opening_ytd"]==0 and d["debit_ytd"]==0 and d["credit_ytd"]==0 and d["closing_ytd"]==0:
            continue
        drcrs_rows.append({"Cat1": v, "Cat2": d["opening"], "Cat3": d["debit"], "Cat4": d["credit"], "Cat5": d["closing"], "Cat6": "", "Cat7": d["opening_ytd"], "Cat8": d["debit_ytd"], "Cat9": d["credit_ytd"], "Cat10": d["closing_ytd"]})
        toc+=d["opening"]; tdc+=d["debit"]; tcc+=d["credit"]; tclc+=d["closing"]
        tocy+=d["opening_ytd"]; tdcy+=d["debit_ytd"]; tccy+=d["credit_ytd"]; tclcy+=d["closing_ytd"]
        
    drcrs_rows.append({"Cat1": "Grand Total", "Cat2": toc, "Cat3": tdc, "Cat4": tcc, "Cat5": tclc, "Cat6": "", "Cat7": tocy, "Cat8": tdcy, "Cat9": tccy, "Cat10": tclcy})
    drcrs_df = pd.DataFrame(drcrs_rows)
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        pl_df.to_excel(writer, sheet_name='P&L', index=False, startrow=6, header=False)
        drcrs_df.to_excel(writer, sheet_name='Drs_Crs', index=False, startrow=4, header=False)
        
    wb = load_workbook(filepath)
    ws_pl = wb['P&L']
    
    ws_pl.cell(row=1, column=1, value="M J P T & Co LLP").font = Font(bold=True)
    ws_pl.cell(row=2, column=1, value=company_name).font = Font(bold=True)
    ws_pl.cell(row=4, column=1, value="P&L Analysis").font = Font(bold=True)
    ws_pl.cell(row=5, column=1, value=f"For the month of {report_month}").font = Font(bold=True)
    
    dark_blue_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    green_fill = PatternFill(start_color="385D3A", end_color="385D3A", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    
    ws_pl.cell(row=6, column=1, value="Particulars").fill = dark_blue_fill
    ws_pl.cell(row=6, column=1).font = white_font
    
    for i, col_name in enumerate(pl_df.columns):
        if col_name == "Particulars": continue
        c_idx = i + 1
        
        display_name = col_name.replace("_YTD", "")
        if display_name == " ": display_name = ""
        
        c = ws_pl.cell(row=6, column=c_idx, value=display_name)
        is_share_trading_col = any(st in display_name for st in share_trading_verticals)
        if "Total" in display_name or is_share_trading_col:
            c.fill = green_fill
        elif display_name == "":
            c.fill = PatternFill(fill_type=None)
        else:
            c.fill = dark_blue_fill
            
        c.font = white_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_pl.column_dimensions[c.column_letter].width = 15
        
    ws_pl.column_dimensions['A'].width = 30
    
    cm_total_idx = list(pl_df.columns).index("Total (including share trading)") + 1 if "Total (including share trading)" in pl_df.columns else list(pl_df.columns).index("Total (without share trading)") + 1
    mon_cell = ws_pl.cell(row=5, column=cm_total_idx, value=report_month.split(' ')[0] + "'" + report_month.split(' ')[1][-2:])
    mon_cell.font = Font(bold=True)
    mon_cell.alignment = Alignment(horizontal="center")
    
    ytd_start_idx = cm_total_idx + 2
    ytd_title_cell = ws_pl.cell(row=5, column=ytd_start_idx, value=ytd_range)
    ytd_title_cell.font = Font(bold=True)
    
    ytd_total_idx = len(pl_df.columns)
    ytd_title2 = ws_pl.cell(row=5, column=ytd_total_idx, value="YTD'"+report_month.split(' ')[1][-2:])
    ytd_title2.font = Font(bold=True)
    ytd_title2.alignment = Alignment(horizontal="center")
    
    last_row = ws_pl.max_row
    ws_pl.cell(row=last_row+2, column=1, value="Business Vertical").fill = dark_blue_fill
    ws_pl.cell(row=last_row+2, column=1).font = white_font
    ws_pl.cell(row=last_row+2, column=2, value=mon_cell.value).fill = dark_blue_fill
    ws_pl.cell(row=last_row+2, column=2).font = white_font
    ws_pl.cell(row=last_row+2, column=3, value="YTD").fill = dark_blue_fill
    ws_pl.cell(row=last_row+2, column=3).font = white_font
    
    r_idx = last_row + 3
    for v in ordered_verticals:
        cm_val = sum(report_data["verticals"][v].get("indirect_breakdown", {}).values())
        ytd_val = sum(report_data["verticals"][v].get("ytd_indirect_breakdown", {}).values())
        if cm_val != 0 or ytd_val != 0:
            ws_pl.cell(row=r_idx, column=1, value=v)
            ws_pl.cell(row=r_idx, column=2, value=cm_val)
            ws_pl.cell(row=r_idx, column=3, value=ytd_val)
            r_idx += 1
            
    ws_pl.cell(row=r_idx, column=1, value="Total Operating Expenses").font = Font(bold=True)
    ws_pl.cell(row=r_idx, column=2, value=cm_totals_inc_st['tot_ind']).font = Font(bold=True)
    ws_pl.cell(row=r_idx, column=3, value=ytd_totals_inc_st['tot_ind']).font = Font(bold=True)
    
    ws_drcr = wb['Drs_Crs']
    ws_drcr.cell(row=1, column=1, value="M J P T & Co LLP").font = Font(bold=True)
    ws_drcr.cell(row=2, column=1, value=company_name).font = Font(bold=True)
    ws_drcr.cell(row=4, column=1, value="Summary of Sundry Debtors & Sundry Creditors").font = Font(bold=True)
    
    mon = report_month.split(' ')[0] + "'" + report_month.split(' ')[1][-2:]
    ws_drcr.cell(row=3, column=3, value=mon).alignment = Alignment(horizontal="center")
    ws_drcr.merge_cells(start_row=3, start_column=3, end_row=3, end_column=6)
    
    ws_drcr.cell(row=3, column=8, value=ytd_range).alignment = Alignment(horizontal="center")
    ws_drcr.merge_cells(start_row=3, start_column=8, end_row=3, end_column=11)
    
    ws_drcr.insert_cols(1)
    ws_drcr.column_dimensions['B'].width = 25
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
        ws_drcr.column_dimensions[col].width = 15
        
    wb.save(filepath)
    return filename
