import win32com.client as win32
import os
import time
import datetime
import shutil
import re
import bank_rules
import excel_validator
from config import BASE_SHAREPOINT_PATH

# --- SETTINGS ---
SLEEP_SECONDS = 10
HISTORY_FILE = "processed_log.txt"

def log(message):
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}")

def sanitize_filename(filename):
    """Remove characters that Outlook or Windows filesystem might not like."""
    # Remove characters like / \ : * ? " < > |
    sanitized = re.sub(r'[\\/*?:"<>|]', "", filename)
    # Ensure it's not too long (max 200 to be safe)
    if len(sanitized) > 200:
        base, ext = os.path.splitext(sanitized)
        sanitized = base[:190] + ext
    return sanitized

def start_outlook_safely():
    """Starts a session. Prioritizes Redemption (RDOSession) for speed and stability."""
    try:
        # 1. TRY REDEMPTION (The Pro Choice)
        try:
            session = win32.Dispatch("Redemption.RDOSession")
            session.Logon()
            log(" > Connected via REDEMPTION (High Speed Mode).")
            return session, True # True means we are using Redemption
        except Exception as e:
            log(f" > Redemption not available ({e}). Falling back to standard Outlook...")
            
        # 2. FALLBACK TO STANDARD WIN32COM
        try:
            outlook_app = win32.GetActiveObject("Outlook.Application")
            log(" > Connected to your open Outlook window.")
        except:
            log(" > Outlook is not running. Starting it automatically...")
            outlook_app = win32.Dispatch("Outlook.Application")
            time.sleep(3) 
            log(" > Outlook started successfully.")
        
        namespace = outlook_app.GetNamespace("MAPI")
        namespace.Logon() 
        return namespace, False # False means standard mode
    except Exception as e:
        log(f"CRITICAL ERROR: Could not start Outlook: {e}")
        return None, None

def load_processed_ids():
    if not os.path.exists(HISTORY_FILE): return set()
    with open(HISTORY_FILE, "r") as f:
        return set(line.strip() for line in f)

def save_processed_id(message_id):
    with open(HISTORY_FILE, "a") as f:
        f.write(f"{message_id}\n")

def get_pull_selection():
    """
    Asks the user which returns they want to scan for.
    Returns: 'BSD2_3', 'BSD4', or 'ALL'
    """
    print("\n--- RETURN TYPE SELECTION ---")
    print("1. Scan for BSD 2 & 3 Returns Only (STRICT)")
    print("2. Scan for BSD 4 Returns Only (STRICT)")
    print("3. Scan for ALL Returns (BSD 2, 3 & 4)")
    
    while True:
        choice = input("Select return type to pull (1, 2, or 3): ").strip()
        if choice == '1': return 'BSD2_3'
        if choice == '2': return 'BSD4'
        if choice == '3': return 'ALL'
        print("Invalid choice. Please select 1, 2, or 3.")

def is_safe_attachment(att, is_redemption=False):
    """Checks if an attachment is safe to process based on extension and hidden status."""
    try:
        fname = att.FileName.lower()
        ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.csv', '.xlsm', '.zip', '.rar', '.msg'}
        
        # 1. EXTENSION CHECK
        if any(fname.endswith(ext) for ext in ALLOWED_EXTENSIONS):
            # Anti-malware: Block double extensions with executable types
            parts = fname.split('.')
            if len(parts) > 2:
                DANGEROUS = {'exe', 'vbs', 'js', 'bat', 'scr', 'msi'}
                if parts[-2] in DANGEROUS: return False
            return True
        
        # 2. EMBEDDED ITEM CHECK (Type 5)
        # Handle Redemption vs Standard Outlook properties
        try:
            att_type = att.Type
            if att_type == 5 and "." not in fname: 
                return True
            
            # Redemption specific: Skip hidden items (system/signature icons)
            if is_redemption and att.Hidden:
                return False
        except:
            pass

        return False
    except:
        return False

def process_message_recursive(message, outlook_ns, target_date, target_type='ALL', is_redemption=False):
    """
    Recursively scans emails and attachments.
    target_type: 'BSD2_3', 'BSD4', or 'ALL'
    """
    successfully_saved_ids = []
    
    try:
        # 1. RECURSION FOR FORWARDS
        for i in range(1, message.Attachments.Count + 1):
            try:
                att = message.Attachments.Item(i)
                fname = att.FileName.lower()
                
                # Redemption RDOAttachment.Type or Outlook Attachment.Type
                try:
                    att_type = att.Type
                except:
                    att_type = 0
                
                if fname.endswith('.msg') or att_type == 5:
                    temp_msg = os.path.join(os.environ['TEMP'], f"temp_{int(time.time())}_{i}.msg")
                    try:
                        att.SaveAsFile(temp_msg)
                        
                        # Inner message extraction
                        if is_redemption:
                            # Redemption uses direct file path for Msg files
                            inner_msg = outlook_ns.GetMessageFromMsgFile(temp_msg)
                        else:
                            inner_msg = outlook_ns.OpenSharedItem(temp_msg)
                            
                        successfully_saved_ids.extend(process_message_recursive(inner_msg, outlook_ns, target_date, target_type, is_redemption))
                        
                        if not is_redemption: inner_msg.Close(0)
                        if os.path.exists(temp_msg): os.remove(temp_msg)
                    except:
                        if os.path.exists(temp_msg): os.remove(temp_msg)
                
            except: 
                continue # Skip attachments we can't even touch

        # 2. ATTACHMENT PROCESSING
        for i in range(1, message.Attachments.Count + 1):
            try:
                att = message.Attachments.Item(i)
                fname = att.FileName
                fname_lower = fname.lower()
            except:
                continue
            
            if not is_safe_attachment(att, is_redemption):
                continue
            
            if not fname_lower.endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                continue

            temp_path = os.path.join(os.environ['TEMP'], f"peek_{int(time.time())}_{sanitize_filename(fname)}")
            try:
                att.SaveAsFile(temp_path)
            except Exception as e:
                log(f"   ! Error saving {fname}: {e}")
                continue
            
            # Regulatory Classification & Verification
            analysis = excel_validator.analyze_return(temp_path, user_selected_mode=target_type)
            
            # Audit Log: Regulatory Format
            log(f"   RETURN_TYPE: {analysis['type']}")
            if analysis.get('sheets'):
                log("   SHEETS:")
                for s in analysis['sheets']:
                    log(f"     - {s}")
            
            if analysis['status'] == 'ACCEPTED':
                log(f"   STATUS: ACCEPTED ({analysis['reason']})")
                
                # Show keywords for BSD4 if they exist
                scores = analysis.get('scores', {})
                matches = analysis.get('matches', {})
                if analysis['type'] == 'BSD4' and matches.get('BSD4'):
                    log(f"   [AUDIT] BSD4 Keywords: {', '.join(matches['BSD4'])}")
            else:
                log(f"   STATUS: REJECTED")
                log(f"   REASON: {analysis['reason']}")
                if os.path.exists(temp_path): os.remove(temp_path)
                continue

            # Handle Structural Conflict (User verification mode)
            if analysis.get('conflict'):
                log(f"   [REJECTED] Structural Mismatch: {analysis['conflict']}")
                if os.path.exists(temp_path): os.remove(temp_path)
                continue

            # Identify Bank (Still used for routing, but structural classification is primary)
            bank_id = analysis.get('bank_name')
            if not bank_id:
                match_rule = bank_rules.get_matching_rule(message, analysis['type'])
                if match_rule:
                    bank_id = match_rule['id']
            
            if not bank_id:
                from bank_rules import BANK_RULES
                fname_upper = fname.upper()
                for rule in BANK_RULES:
                    if any(n in fname_upper for n in rule.get('names', [])):
                        bank_id = rule['id']
                        break
            
            if not bank_id: bank_id = "UNKNOWN_BANK"

            # ROUTING
            target_subfolders = []
            if analysis['type'] == 'BSD2_3':
                target_subfolders.append("BSD2_3 Returns")
            elif analysis['type'] == 'BSD4':
                target_subfolders.append("BSD4 Returns")

            # ZB SPLIT
            save_bank_ids = [bank_id]
            if bank_id == "ZB_GROUP": save_bank_ids = ["ZBBANK", "ZBBS"]

            for subfolder in target_subfolders:
                for b_id in save_bank_ids:
                    ext = os.path.splitext(fname_lower)[1]
                    clean_name = f"{b_id}{ext}" if subfolder == "BSD2_3 Returns" else f"BSD4_{b_id}{ext}"
                    
                    final_date = analysis['date'] if analysis['date'] else target_date
                    save_dir = os.path.join(BASE_SHAREPOINT_PATH, subfolder, final_date.strftime("%d-%m-%Y"))
                    os.makedirs(save_dir, exist_ok=True)
                    
                    save_path = os.path.join(save_dir, clean_name)
                    if os.path.exists(save_path):
                        v = 2
                        while os.path.exists(os.path.join(save_dir, f"{os.path.splitext(clean_name)[0]}_v{v}{ext}")): v += 1
                        save_path = os.path.join(save_dir, f"{os.path.splitext(clean_name)[0]}_v{v}{ext}")
                    
                    try:
                        shutil.copy2(temp_path, save_path)
                        report_id = f"BSD4_{b_id}" if subfolder == "BSD4 Returns" else b_id
                        successfully_saved_ids.append(report_id)
                        log(f"   \u2713 Saved to {subfolder}: {os.path.basename(save_path)}")
                    except Exception as e:
                        log(f"   ! Error saving to {subfolder}: {e}")

            if os.path.exists(temp_path): os.remove(temp_path)

    except Exception as e:
        log(f"   [!] Error: {e}")
        
    return successfully_saved_ids

def run_persistent_bot():
    log(f"--- BOT ONLINE: INTELLIGENT ROUTER MODE ---")
    target_type = get_pull_selection()
    log(f" > Monitoring for: {target_type} Returns")
    
    outlook_ns, is_redemption = start_outlook_safely()
    if not outlook_ns: return

    TARGET_EMAIL = "schinogara@rbz.co.zw"
    inbox = None
    
    # Redemption RDOSession.Stores vs Namespace.Folders
    folders_to_scan = outlook_ns.Stores if is_redemption else outlook_ns.Folders
    for folder in folders_to_scan:
        if TARGET_EMAIL.lower() in folder.Name.lower():
            inbox = folder.GetDefaultFolder(6) if is_redemption else folder.Folders("Inbox")
            break
    
    if not inbox:
        log(f"Account {TARGET_EMAIL} not found. Select manually:")
        folders_list = [f for f in folders_to_scan]
        for i, f in enumerate(folders_list): print(f"{i+1}. {f.Name}")
        choice = int(input("Select account number: "))
        selected_account = folders_list[choice-1]
        inbox = selected_account.GetDefaultFolder(6) if is_redemption else selected_account.Folders("Inbox")

    try:
        while True:
            processed_ids = load_processed_ids()
            messages = inbox.Items
            # Redemption sorting fix: ensure the field name is a clean string
            if is_redemption:
                messages.Sort("ReceivedTime", True)
            else:
                messages.Sort("[ReceivedTime]", True)
            
            # Check latest 20 items
            for message in list(messages)[:20]:
                try:
                    if message.EntryID in processed_ids: continue
                    default_date = datetime.date.today() 
                    saved_ids = process_message_recursive(message, outlook_ns, default_date, target_type, is_redemption)
                    if saved_ids:
                        save_processed_id(message.EntryID)
                except: continue
            
            log("... Waiting ...")
            time.sleep(SLEEP_SECONDS)
    except KeyboardInterrupt: pass

def run_historical_test():
    log(f"--- HISTORICAL SCAN: INTELLIGENT ROUTER ---")
    target_type = get_pull_selection()
    
    outlook_ns, is_redemption = start_outlook_safely()
    if not outlook_ns: return

    folders_to_scan = outlook_ns.Stores if is_redemption else outlook_ns.Folders
    folders_list = [f for f in folders_to_scan]
    for i, f in enumerate(folders_list): print(f"{i+1}. {f.Name}")
    choice = int(input("Select account number: "))
    selected_account = folders_list[choice-1]
    inbox = selected_account.GetDefaultFolder(6) if is_redemption else selected_account.Folders("Inbox")

    date_str = input("Enter date (DD-MM-YYYY): ")
    day, month, year = map(int, date_str.split('-'))
    target_date = datetime.date(year, month, day)

    log(f" > Scanning Inbox for {target_date}...")
    try:
        messages = inbox.Items
        if is_redemption:
            messages.Sort("ReceivedTime", True)
        else:
            messages.Sort("[ReceivedTime]", True)
    except Exception as e:
        log(f"CRITICAL: Lost connection to Outlook: {e}")
        return
    
    bsd23_submitted = set()
    bsd4_submitted = set()
    total_found = 0

    try:
        msg_list = list(messages)
    except Exception as e:
        log(f"CRITICAL: Failed to load messages: {e}")
        return

    for message in msg_list:
        try:
            _ = message.Subject 
            msg_date = message.ReceivedTime.date()
            if msg_date > target_date: continue
            if msg_date < target_date: break
            
            log(f"   [CHECKING] {message.Subject}")
            saved_ids = process_message_recursive(message, outlook_ns, target_date, target_type, is_redemption)
            if saved_ids:
                total_found += len(saved_ids)
                for b_id in saved_ids:
                    if "ZB_GROUP" in b_id or "ZBBANK" in b_id or "ZBBS" in b_id:
                         target_sets = []
                         if "BSD4_" in b_id:
                             if target_type in ['ALL', 'BSD4']: target_sets.append(bsd4_submitted)
                         else:
                             if target_type in ['ALL', 'BSD2_3']: target_sets.append(bsd23_submitted)
                         for s in target_sets:
                             s.add("ZBBANK")
                             s.add("ZBBS")
                    elif b_id.startswith("BSD4_"):
                        if target_type in ['ALL', 'BSD4']: bsd4_submitted.add(b_id.replace("BSD4_", ""))
                    else:
                        if target_type in ['ALL', 'BSD2_3']: bsd23_submitted.add(b_id)
        except Exception as e:
            if "remote procedure call failed" in str(e).lower():
                log("!!! Outlook connection lost.")
                break
            continue

    log("\n" + "="*60)
    log(f"SUBMISSION REPORT: {target_date.strftime('%d %B %Y')}")
    log("="*60)

    def print_section(title, submitted_set, expected_list):
        log(f"\n--- {title} ---")
        log(f"Collected: {len(submitted_set)} / {len(expected_list)}")
        if submitted_set:
            log("\u2713 SUBMITTED:")
            for bank in sorted(submitted_set): log(f"   \u2022 {bank}")
        missing = set(expected_list) - submitted_set
        if missing:
            log("\u2717 MISSING:")
            for bank in sorted(missing): log(f"   \u2022 {bank}")

    if target_type in ['ALL', 'BSD2_3']: print_section("BSD 2 & 3 RETURNS", bsd23_submitted, bank_rules.BSD23_EXPECTED_BANKS)
    if target_type in ['ALL', 'BSD4']: print_section("BSD 4 RETURNS", bsd4_submitted, bank_rules.BSD4_EXPECTED_BANKS)
    
    log("\n" + "="*60)
    os.system("pause")

if __name__ == "__main__":
    print("--- RBZ BSD AUTOMATION BOT ---")
    print("MODE: INTELLIGENT ROUTER (Auto-Detect BSD 2,3,4)")
    
    print("\n1. Live Monitor")
    print("2. Historical Scan")
    
    mode_choice = input("Select mode (1 or 2): ").strip()
    
    if mode_choice == "2":
        run_historical_test()
    else:
        run_persistent_bot()
