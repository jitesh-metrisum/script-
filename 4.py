import os
from datetime import datetime  # Added import
from playwright.sync_api import Playwright, sync_playwright

# 📂 Base path where all reports are saved (your Desktop → A folder)
BASE_DIR = r"C:\Users\maa00\OneDrive - Hem Corporation Pvt Ltd\Dashboard Sources\Focus Files"

# 🔑 Credentials
USERNAME = "MIS Local"
PASSWORD = "Mis@123"
COMPANY = "180"
URL = "http://13.235.198.88/focus8w/"


def login(page):
    """Reusable login"""
    print("🔐 Logging in...")
    page.goto(URL, timeout=0)

    page.wait_for_selector("#txtUsername", timeout=10000)
    page.fill("#txtUsername", USERNAME)

    page.wait_for_selector("#txtPassword", timeout=10000)
    page.fill("#txtPassword", PASSWORD)

    page.select_option("#ddlCompany", value=COMPANY)
    page.click("#btnSignin")
    page.wait_for_timeout(2000)
    print("✅ Logged in.")


def download_report(page, folder, filename=None):
    """Handles download and saves to correct folder with a date stamp."""
    os.makedirs(folder, exist_ok=True)
    
    print("⏳ Starting download...")
    with page.expect_download(timeout=30 * 60 * 1000) as download_info:
        page.locator("#reportViewControls span:has-text('Ok')").click()

    download = download_info.value
    
    # Determine the original filename
    original_filename = filename or download.suggested_filename
    
    # Get the current date as a string (e.g., "2025-09-18")
    date_str = datetime.now().strftime("%Y-%m-%d")
    
    # Split the filename from its extension
    base_name, extension = os.path.splitext(original_filename)
    
    # Create the new, unique filename
    new_filename = f"{base_name}_{date_str}{extension}"
    
    # Set the final save path
    save_path = os.path.join(folder, new_filename)
    
    download.save_as(save_path)
    print(f"✅ Saved to: {save_path}")


# ------------------------- REPORT FUNCTIONS ------------------------- #

def run_closing_stock_fg(page):
    folder = os.path.join(BASE_DIR, "Closing Stock FG")
    print("📊 Closing Stock FG...")
    page.fill("#txtSearchMenu_MainLayout", "Closing Stock F G")
    page.keyboard.press("Enter")
    page.wait_for_timeout(2000)
    page.select_option("#RITOutput_", value="3")
    download_report(page, folder)


def run_closing_stock_PPM(page):
    folder = os.path.join(BASE_DIR, "Closing Stock PPM")
    print("📊 Closing Stock PPM...")

    page.wait_for_selector("#txtSearchMenu_MainLayout", timeout=5000)
    page.fill("#txtSearchMenu_MainLayout", "Closing Stock PPM")
    page.keyboard.press("Enter")
    page.wait_for_timeout(1000)
    page.wait_for_selector("#RITOutput_", timeout=5000)
    page.select_option("#RITOutput_", value="3")

    download_report(page, folder, "Closing_Stock_PPM.xlsx")


def run_customer_ageing_gt(page):
    folder = os.path.join(BASE_DIR, "GT Outstanding Reports")
    print("📊 Customer Ageing (General Trade)...")
    page.get_by_role("link", name=" Financials").click()
    page.get_by_role("link", name="Receivable and Payable").click()
    page.get_by_role("link", name="Customer Detail ").click()
    page.get_by_role("link", name="Customer Ageing Details").click()

    page.get_by_text("Sundry Debtors", exact=True).first.click()
    page.get_by_role("cell", name="Sundry Debtors - Domestic", exact=True).dblclick()
    page.get_by_role("cell", name="Area - General Trade", exact=True).click()
    page.get_by_role("row", name="123 Area - General Trade 122122 Customer", exact=True).get_by_label("").check()
    page.get_by_role("checkbox", name="Adjustment as on Today").check()

    page.locator("#RITOutput_").select_option("3")
    page.wait_for_timeout(1000)
    page.locator("#RITLayout_").select_option("1537")
    
    download_report(page, folder, "Customer_Ageing_GT.xlsx")


def run_link_chain_analysis(page):
    folder = os.path.join(BASE_DIR, "Link Chain Analysis 25-26")
    print("📊 Link Chain Analysis...")
    page.get_by_role("link", name=" Inventory").click()
    page.get_by_role("link", name="Order Management ").click()
    page.get_by_role("link", name="Analysis of Linked/Unlinked").click()
    page.get_by_role("link", name="Link chain analysis").click()

    page.select_option("#DateOptions_", value="9")
    page.locator("#RITTable__1").fill("g")
    page.get_by_role("cell", name="GT Sales Work Flow", exact=True).click()
    page.locator("#RITTable__1").press("Enter")

    page.select_option("#RITLayout_", value="1474")

    page.select_option("#RITOutput_", value="3")
    download_report(page, folder)

# -----------------------------------------------------------------------------------------
def run_sales_order_export(page):
    folder = os.path.join(BASE_DIR, "Sales Order Report Export")
    print("📑 Running Sales Order Export...")

    page.get_by_role("textbox", name="Menu Search").click()
    page.get_by_role("textbox", name="Menu Search").fill("Sales Order Export")
    page.get_by_role("textbox", name="Menu Search").press("ArrowDown")
    page.locator("#homeMenuRun").get_by_text("Sales Order Export").click()

    page.locator("#RITOutput_").select_option("3")

    download_report(page, folder)



# ------------------------- MAIN ------------------------- #

def main(playwright: Playwright):
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    login(page)

    # Run in exact sequence
    # run_closing_stock_fg(page)
    # run_closing_stock_PPM(page)
    # run_customer_ageing_gt(page)
    run_sales_order_export(page)
    # run_link_chain_analysis(page)

    context.close()
    browser.close()


if __name__ == "__main__":
    with sync_playwright() as playwright:
        main(playwright)
