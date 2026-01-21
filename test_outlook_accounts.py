import win32com.client

print("=== OUTLOOK ACCOUNT DIAGNOSTIC ===")
print("Connecting to Outlook...")

try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook.Logon()
    
    print("\nAccounts found in your Outlook profile:")
    print("-" * 50)
    
    for i, folder in enumerate(outlook.Folders, 1):
        print(f"{i}. Account Name: '{folder.Name}'")
    
    print("-" * 50)
    print("\nIf you see your RBZ account above, the bot should now work!")
    
except Exception as e:
    print(f"\nERROR: {e}")
    print("\nTroubleshooting:")
    print("1. Make sure you're running this in the same terminal (not as Admin)")
    print("2. Check if Outlook is running in Task Manager")
    print("3. Try closing Outlook completely and running this again")

input("\nPress Enter to exit...")
