import openpyxl
import os

def inspect_file(filename):
    try:
        print(f"Inspecting {filename}...")
        wb = openpyxl.load_workbook(filename, data_only=True)
        sheet = wb.active
        output = []
        output.append(f"=== FILE: {filename} ===")
        output.append(f"Sheet: {sheet.title}")
        for r in range(1, 60):
            row_data = []
            for c in range(1, 15): # Look at first 15 columns
                val = sheet.cell(row=r, column=c).value
                if val is not None:
                    row_data.append(str(val))
            if row_data:
                output.append(f"Row {r}: {' | '.join(row_data)}")
        output.append("\n")
        return "\n".join(output)
    except Exception as e:
        return f"Error {filename}: {str(e)}\n"

if __name__ == "__main__":
    content = ""
    if os.path.exists('FIRSTCAPITAL.xlsx'):
        content += inspect_file('FIRSTCAPITAL.xlsx')
    if os.path.exists('IDBZ.xlsx'):
        content += inspect_file('IDBZ.xlsx')
    
    with open('excel_inspection.txt', 'w', encoding='utf-8') as f:
        f.write(content)
    print("Done.")
