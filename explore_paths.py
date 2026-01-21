import os
import config

def log(msg):
    with open("path_debug.txt", "a") as f:
        f.write(msg + "\n")

# Clear log
with open("path_debug.txt", "w") as f:
    f.write("--- PATH DIAGNOSTIC ---\n")

path = config.BASE_SHAREPOINT_PATH
log(f"Config Path: [{path}]")

if os.path.exists(path):
    log("Path EXISTS OK!")
else:
    log("Path NOT FOUND.")
    
    # Check segment by segment
    parts = path.split('\\')
    current = parts[0] + '\\'
    
    for i, part in enumerate(parts[1:]):
        current = os.path.join(current, part)
        exists = os.path.exists(current)
        log(f"[{'OK' if exists else 'FAIL'}] {current}")
        
        if not exists:
            log(f" -> Stopped at: {part}")
            # Try to list parent to see what IS there
            parent = os.path.dirname(current)
            if os.path.exists(parent):
                log(f" -> Contents of {parent}:")
                try:
                    for x in os.listdir(parent):
                        log(f"    - {x}")
                except Exception as e:
                    log(f"    Error listing: {e}")
            break
