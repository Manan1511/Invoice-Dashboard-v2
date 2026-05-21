from typing import List, Dict

class FinancialValidationError(Exception):
    """Custom exception raised for financial data validation errors."""
    def __init__(self, title: str, message: str, details: List[str] = None):
        super().__init__(message)
        self.title = title
        self.details = details or []

def validate_sheets_exist(actual_sheets: List[str], required_sheets: List[str]) -> Dict[str, str]:
    """
    Validates that required sheets exist in the Excel file.
    Uses normalization (stripping spaces) to safely handle trailing spaces like 'TB '.
    Returns a dictionary mapping the normalized required sheet name to the actual sheet name.
    """
    normalized_sheets = {name.strip(): name for name in actual_sheets}
    missing = []
    mapped_sheets = {}
    
    for req in required_sheets:
        req_norm = req.strip()
        if req_norm not in normalized_sheets:
            missing.append(req)
        else:
            mapped_sheets[req_norm] = normalized_sheets[req_norm]
            
    if missing:
        raise FinancialValidationError(
            title="Missing Required Sheets",
            message="The uploaded Excel file is missing one or more required sheets.",
            details=[f"Missing sheet: '{m}'" for m in missing]
        )
    
    return mapped_sheets

def validate_trial_balance(tb_df) -> None:
    """
    Validates that the Trial Balance debits and credits match.
    Assumes columns 'Debit' and 'Credit' exist.
    """
    if 'Debit' not in tb_df.columns or 'Credit' not in tb_df.columns:
        raise FinancialValidationError(
            title="Invalid Trial Balance Format",
            message="The Trial Balance sheet must contain 'Debit' and 'Credit' columns.",
            details=[]
        )
        
    import pandas as pd
    debits = pd.to_numeric(tb_df['Debit'], errors='coerce').fillna(0)
    credits = pd.to_numeric(tb_df['Credit'], errors='coerce').fillna(0)
    
    total_debit = debits.sum()
    total_credit = credits.sum()
    
    # Using round to avoid floating point precision issues
    if round(total_debit, 2) != round(total_credit, 2):
        delta = abs(total_debit - total_credit)
        raise FinancialValidationError(
            title="Unbalanced Trial Balance",
            message="The Trial Balance debits and credits do not match.",
            details=[
                f"Total Debit: {total_debit:,.2f}",
                f"Total Credit: {total_credit:,.2f}",
                f"Variance: {delta:,.2f}"
            ]
        )
