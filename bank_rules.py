
# Expected banks per report type
BSD23_EXPECTED_BANKS = [
    "EMPOWERBANK", "BANCABC", "STANBIC", "FBCBANK", "FBCCROWN", "AFC", "SUCCESS", 
    "TIMEBANK", "GETBUCKS", "STEWARD", "ACL", "NMB", "NBS", "ZWMB",
    "FIRSTCAPITAL", "METBANK", "MUKURU", "INNBUCKS", "CBZ",
    "NEDBANK", "ECOBANK", "IDBZ", "CABS", "ZBBANK", "ZBBS", "POSB"
]

BSD4_EXPECTED_BANKS = [
    "EMPOWERBANK", "BANCABC", "STANBIC", "FBCBANK", "FBCCROWN", "AFC", "SUCCESS", 
    "TIMEBANK", "GETBUCKS", "STEWARD", "ACL", "NMB", "NBS", "ZWMB",
    "FIRSTCAPITAL", "METBANK", "MUKURU", "INNBUCKS", "CBZ",
    "NEDBANK", "ECOBANK", "IDBZ", "CABS", "ZBBANK", "ZBBS", "POSB"
]

# For now they are the same, but the user can adjust BSD4_EXPECTED_BANKS if some banks don't submit BSD4.

DAILY_EXPECTED_BANKS = [
    "AFC", "BANCABC", "FIRSTCAPITAL", "CBZ", "ECOBANK", "FBCBANK",
    "NEDBANK", "METBANK", "NMB", "STANBIC", "FBCCROWN", "STEWARD",
    "ZBBANK", "CABS", "NBS", "ZBBS", "POSB"
]

# Priority Ordered Rules
BANK_RULES = [
    {"id": "TIMEBANK", "sender_email": "ninoymatcheso@gmail.com", "save_as": "TIMEBANK", "names": ["TIMEBANK"]},
    {"id": "FBCBANK", "sender_email": "godwill.dube@fbc.co.zw", "save_as": "FBCBANK", "names": ["FBC BANK"]},
    {"id": "FBCCROWN", "sender_domains": ["@fbc.co.zw"], "subject_includes": ["CROWN"], "save_as": "FBCCROWN", "names": ["CROWN"]},
    {"id": "ZB_GROUP", "sender_domains": ["@zb.co.zw"], "save_as": "ZB_GROUP", "special_handler": "ZB_SPLIT", "names": ["ZB BANK", "ZB BUILDING", "ZBBS", "ZB GROUP"]},
    {"id": "EMPOWERBANK", "sender_domains": ["@empowerbank.co.zw"], "save_as": "EMPOWERBANK", "names": ["EMPOWER"]},
    {"id": "BANCABC", "sender_domains": ["@bancabc.co.zw"], "save_as": "BANCABC", "names": ["BANCABC", "BANC ABC"]},
    {"id": "STANBIC", "sender_domains": ["@stanbic.com"], "save_as": "STANBIC", "names": ["STANBIC"]},
    {"id": "AFC", "sender_domains": ["@afcholdings.co.zw"], "save_as": "AFC", "names": ["AFC", "AGRICULTURAL"]},
    {"id": "SUCCESS", "sender_domains": ["@successbank.co.zw"], "save_as": "SUCCESS", "names": ["SUCCESS"]},
    {"id": "GETBUCKS", "sender_domains": ["@getbucksbank.com"], "save_as": "GETBUCKS", "names": ["GETBUCKS"]},
    {"id": "STEWARD", "sender_domains": ["@stewardbank.co.zw"], "save_as": "STEWARD", "names": ["STEWARD", "TN CYBER", "TN CYBERTECH"]},
    {"id": "ACL", "sender_domains": ["@africancentury.co.zw"], "save_as": "ACL", "names": ["AFRICAN CENTURY", "ACL"]},
    {"id": "NMB", "sender_domains": ["@nmbz.co.zw"], "save_as": "NMB", "names": ["NMB"]},
    {"id": "NBS", "sender_domains": ["@nbs.co.zw"], "save_as": "NBS", "names": ["NBS", "NATIONAL BUILDING"]},
    {"id": "ZWMB", "sender_domains": ["@womensbank.co.zw"], "save_as": "ZWMB", "names": ["WOMEN", "ZWMB"]},
    {"id": "FIRSTCAPITAL", "sender_domains": ["@firstcapitalbank.co.zw"], "save_as": "FIRSTCAPITAL", "names": ["FIRST CAPITAL", "FIRSTCAPITAL"]},
    {"id": "METBANK", "sender_domains": ["@metbank.co.zw"], "save_as": "METBANK", "names": ["METBANK", "MET BANK"]},
    {"id": "MUKURU", "sender_domains": ["@mukuru.com"], "save_as": "MUKURU", "names": ["MUKURU"]},
    {"id": "INNBUCKS", "sender_domains": ["@innbucks.co.zw"], "save_as": "INNBUCKS", "names": ["INNBUCKS"]},
    {"id": "CBZ", "sender_domains": ["@cbz.co.zw"], "save_as": "CBZ", "names": ["CBZ"]},
    {"id": "NEDBANK", "sender_domains": ["@nedbank.co.zw"], "save_as": "NEDBANK", "names": ["NEDBANK"]},
    {"id": "ECOBANK", "sender_domains": ["@ecobank.com"], "save_as": "ECOBANK", "names": ["ECOBANK"]},
    {"id": "IDBZ", "sender_domains": ["@idbz.co.zw"], "save_as": "IDBZ", "names": ["IDBZ"]},
    {"id": "CABS", "sender_domains": ["@oldmutual.co.zw"], "save_as": "CABS", "names": ["CABS"]},
    {"id": "POSB", "sender_domains": ["@posb.co.zw"], "save_as": "POSB", "names": ["POSB"]}
]

def rule_is_allowed(rule_id, allowed_banks):
    if rule_id in allowed_banks:
        return True
    if rule_id == "ZB_GROUP":
        return bool({"ZBBANK", "ZBBS"} & set(allowed_banks))
    return False

def email_domain(sender_email):
    return sender_email.split("@")[-1]

def get_matching_rule(message, report_type='BSD'):
    subject = message.Subject.upper()
    try:
        sender_name = message.SenderName.upper()
        if message.SenderEmailType == "EX":
            sender_email = message.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
        else:
            sender_email = message.SenderEmailAddress.lower()
    except:
        sender_email = getattr(message, "SenderEmailAddress", "").lower()
        sender_name = ""

    if report_type == 'BSD2_3':
        allowed_banks = BSD23_EXPECTED_BANKS
    elif report_type == 'BSD4':
        allowed_banks = BSD4_EXPECTED_BANKS
    else:
        allowed_banks = DAILY_EXPECTED_BANKS

    BSD_KEYWORDS = ["BSD", "BSD2", "BSD3", "RETURN", "STANDARDISED", "STANDARDIZED", "FOREIGN", "EXPOSURE", "SRF", "BSA", "CODED", "MARGIN", "LCR"]

    # 1. Subject-based Filter (BSD vs DAILY Logic) - Now a filter, not a gatekeeper
    IS_DAILY_SUBJECT = "DAILY" in subject and not any(k in subject for k in ["BSD", "RETURN", "EXPOSURE"])
    IS_BSD_SUBJECT = any(k in subject for k in ["BSD", "RETURN", "EXPOSURE", "FOREIGN", "BSA", "SRF", "LCR", "MARGIN", "A1", "MAP", "ERROR"])
    
    # If we are in BSD mode and it's CLEARLY a Daily subject with No BSD signal, skip.
    # if report_type == 'BSD' and IS_DAILY_SUBJECT:
    #     return None

    # 2. Iterate through rules
    for rule in BANK_RULES:
        if not rule_is_allowed(rule["id"], allowed_banks):
            continue

        # log(f"      [DEBUG] Checking Rule: {rule['id']} vs Sender: {sender_email}") # Quiet mode
        match = False
        # A. Exact Email Match
        if "sender_email" in rule and rule["sender_email"] == sender_email:
            match = True
        
        # B. Robust Domain Matching
        elif "sender_domains" in rule:
            sender_dom = email_domain(sender_email)
            if any(sender_dom.endswith(d.replace("@", "")) for d in rule["sender_domains"]):
                match = True
        
        # C. RBZ Forward Detection (Hardened)
        IS_RBZ_FORWARD = ("@rbz.co.zw" in sender_email or "RBZ" in sender_name)
        if not match and IS_RBZ_FORWARD:
            name_check = rule.get("names", [rule["id"]])
            if any(k.upper() in subject for k in name_check):
                match = True

        if not match: continue

        # D. Rule-Specific Refinement (Subject inclusions/exclusions)
        if "subject_includes" in rule:
            if not any(k.upper() in subject for k in rule["subject_includes"]): continue
        if "subject_excludes" in rule:
            if any(k.upper() in subject for k in rule["subject_excludes"]): continue

        return rule

    return None
