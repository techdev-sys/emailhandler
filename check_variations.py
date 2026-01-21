import os

# Common User Base
base = r"C:\Users\chinogs\OneDrive - Reserve Bank of Zimbabwe"

variations = [
    # 1. User provided (Space before 'S)
    r"Banking Supervision, Surveillance & Financial Stability - GOVERNOR 'S DATA REQUESTS\2025__",
    # 2. Correct Grammar (No space)
    r"Banking Supervision, Surveillance & Financial Stability - GOVERNOR'S DATA REQUESTS\2025__", 
    # 3. Case variations or spacing
    r"Banking Supervision, Surveillance & Financial Stability - GOVERNORS DATA REQUESTS\2025__",
    r"Banking Supervision, Surveillance & Financial Stability - GOVERNOR S DATA REQUESTS\2025__",
]

print(f"Checking Base: {base}")
if os.path.exists(base):
    print(" [OK] Base Exists")
    
    # Check deeper
    for v in variations:
        full = os.path.join(base, v)
        if os.path.exists(full):
            print(f" [FOUND!] Match: {full}")
            exit(0)
        else:
             # Check partial
             part1 = v.split('\\')[0]
             if os.path.exists(os.path.join(base, part1)):
                 print(f" [PARTIAL MATCH] {part1} exists, but 2025__ might be missing.")
             else:
                 print(f" [FAIL] {v}")
    
    # If no match, list the base directory to see what IS there
    print("\nListing Base Directory content:")
    try:
        for x in os.listdir(base):
            print(f" - {x}")
    except: pass

else:
    print(" [FAIL] Base does not exist.")
    # Check C:\Users\chinogs
    print("Listing User Home:")
    try:
        for x in os.listdir(r"C:\Users\chinogs"):
            if "OneDrive" in x: print(f" - {x}")
    except: pass
