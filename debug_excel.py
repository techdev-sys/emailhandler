import zipfile
import re
import os

def analyze(filepath):
    print(f"Analyzing {filepath}...")
    if not os.path.exists(filepath):
        print("File does not exist.")
        return
    
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            print("Files in ZIP:", z.namelist())
            if 'xl/sharedStrings.xml' in z.namelist():
                content = z.read('xl/sharedStrings.xml').decode('utf-8', errors='ignore')
                strings = re.findall(r'<t.*?>(.*?)</t>', content)
                print("\nUNIQUE STRINGS (SET):")
                unique_strings = sorted(list(set(s.upper() for s in strings if len(s) < 100)))
                for s in unique_strings[:100]:
                    print(f"  - {s}")
            else:
                print("No sharedStrings.xml found.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    analyze(r"c:\Users\chinogs\Music\RBZ_Auto_Bot\FIRSTCAPITAL.xlsx")
