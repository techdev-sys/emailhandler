import os

files_to_delete = [
    "analyze_sample.py", "analyze_zip.py", "check_content.py", "check_date.py", 
    "check_excel.py", "check_variations.py", "debug_excel.py", "dump_cells.py", 
    "explore_paths.py", "fast_inspect.py", "fast_output.txt", "find_bank_name.py", 
    "find_closest_path.py", "inspect_firstcapital.py", "peek_sample.py", 
    "test_matching.py", "test_outlook_accounts.py", "test_validator_accuracy.py", 
    "test_validator_patch.py", "verify_router.py", "FIRSTCAPITAL.xlsx", "IDBZ.xlsx"
]

for f in files_to_delete:
    if os.path.exists(f):
        try:
            os.remove(f)
            print(f"Deleted {f}")
        except Exception as e:
            print(f"Failed to delete {f}: {e}")
    else:
        print(f"File not found: {f}")
