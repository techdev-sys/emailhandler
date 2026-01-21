
import zipfile
import re

file_path = r"c:\Users\chinogs\Music\RBZ_Auto_Bot\ACL_2026-01-14.xlsx"

try:
    with zipfile.ZipFile(file_path, 'r') as z:
        # sharedStrings.xml contains most labels
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                content = f.read().decode('utf-8')
                # Simple regex to find text between <t> and </t> tags
                strings = re.findall(r'<t>(.*?)</t>', content)
                print("Shared Strings Found (First 20):")
                for s in strings[:20]:
                    print(f"  - {s}")
        else:
            print("sharedStrings.xml not found.")
            
        # Also check sheet names via workbook.xml
        if 'xl/workbook.xml' in z.namelist():
             with z.open('xl/workbook.xml') as f:
                content = f.read().decode('utf-8')
                sheets = re.findall(r'name="(.*?)"', content)
                print("\nSheet Names:")
                for s in sheets:
                    print(f"  - {s}")
except Exception as e:
    print(f"Error: {e}")
