import zipfile
import re
import os

def fast_inspect(filepath):
    output = []
    output.append(f"\n{'='*30}\nFast Inspecting: {filepath}\n{'='*30}")
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            # 1. Shared Strings
            if 'xl/sharedStrings.xml' in z.namelist():
                content = z.read('xl/sharedStrings.xml').decode('utf-8', errors='ignore')
                shared_strings = re.findall(r'<t.*?>(.*?)</t>', content)
                output.append("\n--- SHARED STRINGS (All containing 'BSD', 'ASSET', 'LIAB', 'CURRENCY') ---")
                for i, s in enumerate(shared_strings):
                    supper = s.upper()
                    if any(k in supper for k in ['BSD', 'ASSET', 'LIAB', 'CURRENCY', 'SPOT', 'AGAINST', 'NET']):
                        output.append(f"{i}: {s}")
                
                output.append("\n--- FIRST 20 SHARED STRINGS ---")
                for i, s in enumerate(shared_strings[:20]):
                    output.append(f"{i}: {s}")

            # 2. Workbook.xml (often has sheet names)
            if 'xl/workbook.xml' in z.namelist():
                content = z.read('xl/workbook.xml').decode('utf-8', errors='ignore')
                output.append("\n--- WORKBOOK XML SNIPPET (Sheet names likely here) ---")
                output.append(content[:2000])

    except Exception as e:
        output.append(f"Error: {e}")
    
    return "\n".join(output)

if __name__ == "__main__":
    full_output = ""
    for f in ['FIRSTCAPITAL.xlsx', 'IDBZ.xlsx']:
        if os.path.exists(f):
            full_output += fast_inspect(f)
        else:
            full_output += f"\nFile not found: {f}\n"

    with open('fast_output.txt', 'w', encoding='utf-8') as f:
        f.write(full_output)
    print("Done.")
