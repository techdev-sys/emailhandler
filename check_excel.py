
import openpyxl
import os

try:
    wb = openpyxl.load_workbook('FIRSTCAPITAL.xlsx', data_only=True)
    sheet = wb.active
    with open('excel_report.txt', 'w') as f:
        f.write(f"Sheet Name: {sheet.title}\n")
        # Read a 10x20 grid
        for r in range(1, 21):
            row_vals = [str(sheet.cell(row=r, column=c).value) for c in range(1, 11)]
            f.write(f"Line {r}: {' | '.join(row_vals)}\n")
except Exception as e:
    with open('excel_report.txt', 'w') as f:
        f.write(f"Error: {str(e)}")
