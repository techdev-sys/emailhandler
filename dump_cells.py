import zipfile
import re
import os

def get_cells(filepath):
    results = {}
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            # Get Shared Strings
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                xml_content = z.read('xl/sharedStrings.xml').decode('utf-8', errors='ignore')
                shared_strings = re.findall(r'<t.*?>(.*?)</t>', xml_content)
            
            # Get Sheet 1 (usually rId1 or sheet1.xml)
            # We try sheet1.xml first
            sheet_path = 'xl/worksheets/sheet1.xml'
            if sheet_path not in z.namelist():
                # fallback
                if 'xl/worksheets/sheet2.xml' in z.namelist():
                    sheet_path = 'xl/worksheets/sheet2.xml'

            if sheet_path in z.namelist():
                xml_content = z.read(sheet_path).decode('utf-8', errors='ignore')
                # Extraction for A1 to F30
                for row in range(1, 40):
                    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                        ref = f"{col}{row}"
                        # Match <c r="A1" ...
                        # Type s (shared string)
                        pattern_s = rf'<c r="{ref}"[^>]*t="s"[^>]*><v>(.*?)</v>'
                        m = re.search(pattern_s, xml_content)
                        if m:
                            idx = int(m.group(1))
                            val = shared_strings[idx] if idx < len(shared_strings) else "???"
                            results[ref] = val
                        else:
                            # Inline string or number
                            # Inline string <is><t>...</t></is>
                            pattern_is = rf'<c r="{ref}"[^>]*><is><t>(.*?)</t></is>'
                            m2 = re.search(pattern_is, xml_content)
                            if m2:
                                results[ref] = m2.group(1)
                            else:
                                # Number <v>...</v> (no t="s") - crude check
                                # This regex is tricky if attributes vary.
                                # Simplified: <c r="A1"><v>123</v></c>
                                pattern_n = rf'<c r="{ref}"[^>]*><v>(.*?)</v>'
                                m3 = re.search(pattern_n, xml_content)
                                if m3:
                                    # check if it was actually t="s" handled above
                                    if 't="s"' not in m3.group(0): 
                                        results[ref] = m3.group(1)
    except Exception as e:
        results["ERROR"] = str(e)
    return results

if __name__ == "__main__":
    with open('cell_dump.txt', 'w', encoding='utf-8') as f:
        for fname in ['FIRSTCAPITAL.xlsx', 'IDBZ.xlsx']:
            if os.path.exists(fname):
                f.write(f"\n=== FILE: {fname} ===\n")
                cells = get_cells(fname)
                if not cells:
                    f.write("No cells found or error.\n")
                for row in range(1, 40):
                    row_vals = []
                    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                        ref = f"{col}{row}"
                        if ref in cells:
                            row_vals.append(f"{col}{row}:{cells[ref]}")
                    if row_vals:
                        f.write(" | ".join(row_vals) + "\n")
            else:
                 f.write(f"\n=== FILE: {fname} NOT FOUND ===\n")
    print("Done.")
