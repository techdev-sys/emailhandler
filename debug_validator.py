import excel_validator
import os

files = ['FIRSTCAPITAL.xlsx', 'IDBZ.xlsx']
for f in files:
    path = os.path.abspath(f)
    print(f"\nAnalyzing: {f}")
    res = excel_validator.analyze_return(path)
    print(f"Result: {res}")
