import sys
import os
import time
import re
import subprocess
import urllib.request
from playwright.sync_api import sync_playwright

def is_chrome_running():
    try:
        urllib.request.urlopen("http://127.0.0.1:9222/json", timeout=1)
        return True
    except Exception:
        return False

def launch_chrome():
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
    ]
    
    chrome_path = None
    for path in chrome_paths:
        if os.path.exists(path):
            chrome_path = path
            break
            
    if not chrome_path:
        print("Error: Google Chrome was not found in standard locations.", file=sys.stderr)
        sys.exit(1)
        
    url = "https://onecomhelp.zendesk.com/explore/dashboard/D039494B7CFB701C4C4AA46BC5EDE5F4D357830DEAA9EFAB3C6A5F7037E4F1B3/tab/39682872"
    user_data_dir = r"C:\Users\corr10\AppData\Local\Temp\opencode\chrome-profile"
    os.makedirs(user_data_dir, exist_ok=True)
    
    cmd = [
        chrome_path,
        "--remote-debugging-port=9222",
        f"--user-data-dir={user_data_dir}",
        url
    ]
    
    try:
        creationflags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    except AttributeError:
        creationflags = 0
        
    print("Launching Chrome in debug mode...")
    subprocess.Popen(cmd, creationflags=creationflags)
    
    # Wait for Chrome to start and the DevTools endpoint to become active
    for _ in range(15):
        if is_chrome_running():
            print("Chrome is up and running.")
            return
        time.sleep(1)
    print("Warning: Timed out waiting for Chrome to start.", file=sys.stderr)

def find_label_by_text(page, text):
    escaped_text = re.escape(text)
    label_locator = page.locator("label:visible").filter(has_text=re.compile(f"^{escaped_text}$")).first
    if label_locator.count() > 0:
        return label_locator
    return None

def find_div_by_text(page, text):
    escaped_text = re.escape(text)
    div_locator = page.locator("div:visible").filter(has_text=re.compile(f"^{escaped_text}$")).first
    if div_locator.count() > 0:
        return div_locator
    return None

def main():
    if not is_chrome_running():
        launch_chrome()
    else:
        print("Chrome is already running on port 9222.")

    print("Connecting to Chrome...")
    try:
        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            target_page = None
            
            # Find the Zendesk page or wait for it to load
            print("Searching for the Zendesk dashboard page...")
            for attempt in range(15):
                for context in browser.contexts:
                    for page in context.pages:
                        if "zendesk.com" in page.url:
                            target_page = page
                            break
                    if target_page:
                        break
                if target_page:
                    break
                time.sleep(1)
                    
            if not target_page:
                print("Zendesk dashboard tab not found. Navigating to the page...")
                if browser.contexts and browser.contexts[0].pages:
                    target_page = browser.contexts[0].pages[0]
                    target_page.goto("https://onecomhelp.zendesk.com/explore/dashboard/D039494B7CFB701C4C4AA46BC5EDE5F4D357830DEAA9EFAB3C6A5F7037E4F1B3/tab/39682872")
                else:
                    context = browser.contexts[0] if browser.contexts else browser.new_context()
                    target_page = context.new_page()
                    target_page.goto("https://onecomhelp.zendesk.com/explore/dashboard/D039494B7CFB701C4C4AA46BC5EDE5F4D357830DEAA9EFAB3C6A5F7037E4F1B3/tab/39682872")

            # Pre-flight synchronization: wait for either the Export button, Dropdown, or Modal to be visible
            print("Waiting for page elements to render and stabilize...")
            page_loaded = False
            for attempt in range(30):
                cancel_btn = target_page.locator("button:has-text('Cancel'):visible").first
                excel_option = find_div_by_text(target_page, "Excel")
                main_export_btn = target_page.locator("button.sc-iQZIRz.dJYOnr:has-text('Export'):visible").first
                if main_export_btn.count() == 0:
                    main_export_btn = target_page.locator("button:has-text('Export'):visible").first
                
                if cancel_btn.count() > 0 or excel_option or main_export_btn.count() > 0:
                    print("Page elements stabilized successfully!")
                    page_loaded = True
                    break
                target_page.wait_for_timeout(1000)

            # Precise state checks
            cancel_btn = target_page.locator("button:has-text('Cancel'):visible").first
            modal_is_open = cancel_btn.count() > 0

            print("Checking if the 'Export dashboard' modal is already open...")
            if modal_is_open:
                print("Modal is already open.")
                image_option = find_label_by_text(target_page, "Image")
                modal_indicator = find_label_by_text(target_page, "Image and PDF")
                
                # Check if modal is open but collapsed (Image and PDF is visible, but Image is hidden)
                if modal_indicator is not None and image_option is None:
                    print("Modal is open, but 'Image and PDF' is collapsed. Clicking 'Image and PDF' to expand it...")
                    modal_indicator.click(force=True)
                    target_page.wait_for_timeout(1500)
                    # Re-locate the image option now that it is expanded
                    image_option = find_label_by_text(target_page, "Image")
            else:
                print("Modal is not open. Checking if Export dropdown is open...")
                excel_option = find_div_by_text(target_page, "Excel")
                dropdown_is_open = excel_option is not None

                if not dropdown_is_open:
                    print("Dropdown is not open. Clicking the main header 'Export' button...")
                    main_export_btn = target_page.locator("button.sc-iQZIRz.dJYOnr:has-text('Export'):visible").first
                    if main_export_btn.count() == 0:
                        main_export_btn = target_page.locator("button:has-text('Export'):visible").first
                    
                    if main_export_btn.count() > 0:
                        main_export_btn.click(force=True)
                        print("Clicked main Export button. Waiting for dropdown...")
                        target_page.wait_for_selector("text=Excel", state="visible", timeout=15000)
                        excel_option = find_div_by_text(target_page, "Excel")
                    else:
                        print("Error: Could not locate the main header 'Export' button.")
                        return

                if excel_option:
                    print("Clicking 'Excel' option in dropdown to open the format modal...")
                    excel_option.click(force=True)
                    print("Waiting for the modal to open...")
                    # Wait for the Cancel button (modal indicator) to become visible
                    target_page.wait_for_selector("button:has-text('Cancel'):visible", state="visible", timeout=15000)
                    image_option = find_label_by_text(target_page, "Image")
                else:
                    print("Error: Could not locate 'Excel' option in dropdown.")
                    return

            # Click the 'Image' checkbox label ONLY if it is currently checked (Idempotent)
            if image_option:
                input_id = image_option.get_attribute("for")
                is_checked = False
                if input_id:
                    input_locator = target_page.locator(f'[id="{input_id}"]')
                    if input_locator.count() > 0:
                        is_checked = input_locator.is_checked()
                
                print(f"Image checkbox checked status: {is_checked}")
                if is_checked:
                    print("Clicking the 'Image' checkbox label to toggle/uncheck it...")
                    image_option.click(force=True)
                    target_page.wait_for_timeout(1500)
                else:
                    print("Image checkbox is already unchecked. Skipping click.")
            else:
                print("Error: Could not locate 'Image' checkbox option.")

            # Capture screenshot to verify
            screenshot_path = os.path.join(os.getcwd(), "screenshot.png")
            print("Taking screenshot of the final state...")
            target_page.screenshot(path=screenshot_path)
            print(f"Screenshot successfully saved to {screenshot_path}")
            
    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)

if __name__ == "__main__":
    main()
