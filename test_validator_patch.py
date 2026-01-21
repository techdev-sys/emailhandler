import excel_validator
import os

def test_filenames():
    # Test cases that should pass based on filename (simulate file not existing or empty)
    # Note: validator tries to open zip. If file doesn't exist, it goes to except block.
    
    cases = [
        "Test_BSD2.xlsx",
        "Test_BSD3.xlsx",
        "BSD 2 Return.xlsx",
        "My_BSD2_Return.xlsx"
    ]
    
    print("Testing Filename Fallback (Non-existent files -> Except block):")
    for name in cases:
        try:
            res = excel_validator.is_valid_bsd_return(name)
            print(f"[{'PASS' if res else 'FAIL'}] {name} -> {res}")
        except Exception as e:
            print(f"[ERROR] {name} -> {e}")

    print("\nTesting known Bad names:")
    bad_cases = [
        "Daily_Return.xlsx",
        "Loans.xlsx"
    ]
    for name in bad_cases:
        res = excel_validator.is_valid_bsd_return(name)
        print(f"[{'PASS' if not res else 'FAIL'}] {name} -> {res}")

if __name__ == "__main__":
    test_filenames()
