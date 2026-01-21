import zipfile
import os

file_path = r"c:\Users\chinogs\Music\RBZ_Auto_Bot\FIRSTCAPITAL.xlsx"

try:
    with zipfile.ZipFile(file_path, 'r') as z:
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                content = f.read().decode('utf-8', errors='ignore').upper()
                print(f"RESERVE BANK in content: {'RESERVE BANK' in content}")
                print(f"BSD in content: {'BSD' in content}")
                print(f"SPOT TRANSACTIONS in content: {'SPOT TRANSACTIONS' in content}")
                print(f"FORWARD TRANSACTIONS in content: {'FORWARD TRANSACTIONS' in content}")
except Exception as e:
    print(f"Error: {e}")
