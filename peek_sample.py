import zipfile
import re
import os

file_path = r"c:\Users\chinogs\Music\RBZ_Auto_Bot\IDBZ.xlsx"

print(f"Analyzing {file_path} (BSD 4 Sample)...")
try:
    with zipfile.ZipFile(file_path, 'r') as z:
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                content = f.read().decode('utf-8', errors='ignore')
                # Find all text between <t> and </t>
                strings = re.findall(r'<t.*?>(.*?)</t>', content)
                print("\n--- SHARED STRINGS (Sample) ---")
                # Print first 60 strings to find headers/B1/unique markers
                for i, s in enumerate(strings[:60]):
                    print(f"  {i+1}: {s}")
        
        if 'xl/workbook.xml' in z.namelist():
             with z.open('xl/workbook.xml') as f:
                content = f.read().decode('utf-8', errors='ignore')
                sheets = re.findall(r'name="(.*?)"', content)
                print("\n--- SHEET NAMES ---")
                for s in sheets:
                    print(f"  - {s}")
except Exception as e:
    print(f"Error: {e}")
