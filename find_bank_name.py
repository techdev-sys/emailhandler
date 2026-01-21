import openpyxl
import re

EXPECTED_BANKS = [
    "EMPOWERBANK", "BANCABC", "STANBIC", "FBCCROWN", "AFC", "SUCCESS", 
    "TIMEBANK", "GETBUCKS", "STEWARD", "ACL", "NMB", "NBS", "ZWMB",
    "FIRSTCAPITAL", "METBANK", "MUKURU", "INNBUCKS", "CBZ",
    "NEDBANK", "ECOBANK", "IDBZ", "CABS", "ZBBANK", "ZBBS", "POSB", "FBCBS"
]

def find_bank_name(filepath):
    print(f"\nScanning {filepath} for Bank Name...")
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
        
        # Scan Top Area (Rows 1-15, Cols 1-10)
        found_name = None
        
        for r in range(1, 16):
            row_vals = []
            for c in range(1, 11):
                val = sheet.cell(row=r, column=c).value
                if val:
                    s_val = str(val).upper()
                    row_vals.append(s_val)
                    
                    # Direct Match in Cell
                    for bank in EXPECTED_BANKS:
                        if bank in s_val:
                            print(f"[MATCH] Found '{bank}' in cell ({r},{c}): {s_val}")
                            found_name = bank
            
            # Context Match: "Name of Institution: X"
            row_str = " | ".join(row_vals)
            if "INSTITUTION" in row_str or "NAME OF BANK" in row_str:
                print(f"[CONTEXT] Found Header Label at Row {r}: {row_str}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    find_bank_name('FIRSTCAPITAL.xlsx')
    find_bank_name('IDBZ.xlsx')
