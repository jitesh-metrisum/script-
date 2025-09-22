import os
from playwright.sync_api import Playwright, sync_playwright

# 📂 Directory where reports will be saved
DOWNLOAD_DIR = r"C:\Users\maa00\OneDrive - Hem Corporation Pvt Ltd\Dashboard Sources\Focus Files\Sales Order Report Export"

def run(playwright: Playwright) -> None:
    # Ensure the folder exists
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    # Open login page
    page.goto("http://13.235.198.88/focus8w/")

    # 🔐 Login
    page.wait_for_selector("#txtUsername", timeout=10000)
    page.fill("#txtUsername", "MIS Local")

    print("🔐 Filling password...")
    page.wait_for_selector("#txtPassword", timeout=10000)
    page.fill("#txtPassword", "Mis@123")  # ⚠️ fill real password 

    page.locator("#ddlCompany").select_option("180")
    page.get_by_role("button", name="Sign In").click()

    # 🔍 Search and open report
    page.get_by_role("textbox", name="Menu Search").click()
    page.get_by_role("textbox", name="Menu Search").fill("Sales Order Export")
    page.get_by_role("textbox", name="Menu Search").press("ArrowDown")
    page.locator("#homeMenuRun").get_by_text("Sales Order Export").click()

    page.locator("#RITOutput_").select_option("3")


    with page.expect_download(timeout=0) as download_info:  # 0 = wait indefinitely
        page.locator("#reportViewControls span:has-text('Ok')").first.click()

    download = download_info.value
    file_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
    download.save_as(file_path)
    print(f"✅ Report downloaded to: {file_path}")


    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
