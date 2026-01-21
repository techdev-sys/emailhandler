import os

# Start from User Home
home = os.path.expanduser("~")
print(f"Scanning from: {home}")

target_phrases = ["Reserve Bank", "Banking Supervision"]

found_roots = []

try:
    # Shallow scan of home dir
    for name in os.listdir(home):
        full = os.path.join(home, name)
        if os.path.isdir(full):
            if "OneDrive" in name or "Reserve Bank" in name:
                print(f"[POTENTIAL ROOT] {full}")
                found_roots.append(full)

    # Dig deeper into potential roots
    for root in found_roots:
        print(f"\nScanning inside: {root}")
        try:
            for root_dir, dirs, files in os.walk(root):
                # Don't go too deep
                depth = root_dir[len(root):].count(os.sep)
                if depth > 3: continue
                
                print(f"  DIR: {root_dir}")
                if "GOVERNOR" in root_dir.upper():
                    print(f"  >>> FOUND GOVERNOR FOLDER: {root_dir}")
        except Exception as e:
            print(f"  Error: {e}")

except Exception as e:
    print(f"Root Error: {e}")
