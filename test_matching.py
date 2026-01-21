
import bank_rules

class MockMsg:
    def __init__(self, subject, sender_email, sender_name):
        self.Subject = subject
        self.SenderEmailAddress = sender_email
        self.SenderName = sender_name
        self.SenderEmailType = "SMTP"

def test_match(subj, email, name):
    msg = MockMsg(subj, email, name)
    rule = bank_rules.get_matching_rule(msg, report_type='BSD')
    print(f"Subject: {subj}")
    print(f"Sender: {name} <{email}>")
    print(f"RESULT: {rule['id'] if rule else 'NONE'}")
    print("-" * 30)

test_match("FW: AFC Commercial Bank BSD 2 & BSD 3 AS AT 16/01/2026", "afcreturns@afcholdings.co.zw", "AFCCompliance Returns")
test_match("FW: AFC Commercial Bank BSD 2 & BSD 3 AS AT 16/01/2026", "SChinogara@rbz.co.zw", "Chinogara, Shepherd")
test_match("FW: ZWMB Foreign Currency Exposure Returns BSD2 & BSD 3 - 16.1.2026", "SChinogara@rbz.co.zw", "Chinogara, Shepherd")
test_match("FW: NMB BANK BSD2 & BSD3 19012026", "SChinogara@rbz.co.zw", "Chinogara, Shepherd")
test_match("FW: BancABC FORM BSD2&3 as at 16-January-2026", "SChinogara@rbz.co.zw", "Chinogara, Shepherd")
