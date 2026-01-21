import excel_validator
import os
import datetime

def verify(filename):
    print(f"\nScanning: {filename}")
    if not os.path.exists(filename):
        print("File not found.")
        return

    try:
        result = excel_validator.analyze_return(filename)
        print(f"  [RESULT] Type: {result['type']}")
        print(f"  [RESULT] Date: {result['date']}")
        print(f"  [RESULT] Bank: {result['bank_name']}")
        print(f"  [RESULT] Confidence: {result['confidence']}")
    except Exception as e:
        print(f"  [ERROR] {e}")

if __name__ == "__main__":
    verify('FIRSTCAPITAL.xlsx')
    verify('IDBZ.xlsx')
