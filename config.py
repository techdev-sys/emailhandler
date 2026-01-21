import datetime
import os

# ==========================================
# USER SETTINGS (CHANGE THESE PATHS)
# ==========================================
# This is where your SharePoint is synced on your PC.
# Tip: Open File Explorer, go to the folder, click the address bar, and copy.
#BASE_SHAREPOINT_PATH = r"C:\Users\chinogs\OneDrive - Reserve Bank of Zimbabwe\Banking Supervision, Surveillance & Financial Stability - GOVERNOR 'S DATA REQUESTS"
# BASE_SHAREPOINT_PATH = r"C:\Users\chinogs\Music\RBZ_Auto_Bot\RBZ_Returns"
_year = datetime.date.today().year
BASE_SHAREPOINT_PATH = rf"C:\Users\chinogs\OneDrive - Reserve Bank of Zimbabwe\Banking Supervision, Surveillance & Financial Stability - GOVERNOR 'S DATA REQUESTS\{_year}"

# ==========================================
# DATE LOGIC ENGINE
# ==========================================
# (You can add specific date offsets here if needed in the future)