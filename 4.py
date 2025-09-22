import os
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright

# --- CHANGE THESE ---
# 1. Read credentials securely from environment variables (GitHub Secrets)
USERNAME = os.environ.get("MIS_USERNAME")
PASSWORD = os.environ.get("MIS_PASSWORD")

# 2. This URL MUST be accessible from the public internet
URL = "https://your-publicly-accessible-url.com" 
# --- END CHANGES ---

COMPANY = "180"

# ... your login function remains the same
def login(page):
    """Reusable Login"""
    print("Logging in...")
    page.goto(URL, timeout=60000) # Increased timeout just in case
    page.wait_for_selector(f"#txtUsername", timeout=10000)
    page.fill(f"#txtUsername", USERNAME)
    # ... rest of your script
