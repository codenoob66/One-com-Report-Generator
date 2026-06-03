import subprocess
import os
import sys
import time
import datetime
import urllib.request
import json
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
        
    url = "https://onecomhelp.zendesk.com/explore/dashboard/D039494B7CFB701C4C4AA46BC5EDE5F4D357830DEAA9EFAB3C6A5F7037E4F1B3/tab/39682812"
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
        
    print("Launching Chrome in debug mode for Email tab...")
    subprocess.Popen(cmd, creationflags=creationflags)
    
    # Wait for Chrome to start and the DevTools endpoint to become active
    for _ in range(15):
        if is_chrome_running():
            print("Chrome is up and running.")
            return
        time.sleep(1)
    print("Warning: Timed out waiting for Chrome to start.", file=sys.stderr)

def run_automation():
    today = datetime.date.today()
    print(f"Today's date: {today}")
    
    current_year = today.year
    current_month_abbr = today.strftime("%b")  # e.g., "Jun"
    current_day = today.day
    
    target_header = f"{current_year}-{current_month_abbr}"
    print(f"Targeting date range: {current_month_abbr} 1st, {current_year} to today ({current_month_abbr} {current_day}, {current_year})")
    
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        except Exception as e:
            print(f"Error connecting to Chrome DevTools: {e}", file=sys.stderr)
            return
            
        target_url = "https://onecomhelp.zendesk.com/explore/dashboard/D039494B7CFB701C4C4AA46BC5EDE5F4D357830DEAA9EFAB3C6A5F7037E4F1B3/tab/39682812"
        target_page = None
        
        # Look for the Zendesk dashboard page
        for context in browser.contexts:
            for page in context.pages:
                if target_url in page.url:
                    target_page = page
                    break
            if target_page:
                break
                
        if not target_page:
            print("Zendesk dashboard tab not found. Navigating to the page...")
            if browser.contexts and browser.contexts[0].pages:
                target_page = browser.contexts[0].pages[0]
                target_page.goto(target_url)
            else:
                context = browser.contexts[0] if browser.contexts else browser.new_context()
                target_page = context.new_page()
                target_page.goto(target_url)
                
        print("Connected to Zendesk dashboard (Email). Waiting for page elements to load...")
        try:
            # Wait up to 60 seconds for the time filter widget to be visible and stable
            target_page.wait_for_selector("#bimeTimeFilterWidget-2", state="visible", timeout=60000)
            # Extra buffer time to let the dashboard completely initialize
            target_page.wait_for_timeout(3000)
        except Exception as e:
            print(f"Error: Timed out waiting for Zendesk dashboard to load: {e}", file=sys.stderr)
            return
            
        # 1. Open the 'Time' dropdown
        if target_page.locator(".simple-range-date").count() == 0:
            print("Opening 'Time' filter dropdown...")
            time_button = target_page.locator("#bimeTimeFilterWidget-2")
            if time_button.count() == 0:
                time_button = target_page.locator("a:has-text('Time')")
            if time_button.count() > 0:
                time_button.first.click()
                target_page.wait_for_timeout(2000)
            else:
                print("Could not find or open 'Time' dropdown.", file=sys.stderr)
                return
                
        # 2. Click 'Custom'
        print("Selecting 'Custom' range preset...")
        custom_option = target_page.locator(".simple-range-date div.radio-text:has-text('Custom')")
        if custom_option.count() > 0:
            custom_option.first.click()
            target_page.wait_for_timeout(1500)
        else:
            print("Could not find 'Custom' preset option.", file=sys.stderr)
            return
            
        # 3. Locate calendars
        from_calendar = target_page.locator(".from-calendar").first
        to_calendar = target_page.locator(".to-calendar").first
        
        if from_calendar.count() == 0 or to_calendar.count() == 0:
            print("Calendar widgets not found.", file=sys.stderr)
            return
            
        # 4. Navigate calendars to current year-month
        months_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        
        def parse_header(header_str):
            try:
                parts = header_str.strip().split("-")
                year = int(parts[0])
                month_idx = months_order.index(parts[1])
                return year, month_idx
            except Exception:
                return None, None
                
        target_year = current_year
        target_month_idx = months_order.index(current_month_abbr)
        
        # Navigate From-Calendar
        print("Aligning From-Calendar month...")
        for attempt in range(24):
            header_text = from_calendar.evaluate("el => el.innerText").strip()
            lines = [l.strip() for l in header_text.split("\n") if "-" in l]
            if not lines:
                break
            cal_header = lines[0]
            if target_header in cal_header:
                break
            cal_year, cal_month_idx = parse_header(cal_header)
            if cal_year is None:
                from_calendar.locator(".icon-arrow-right14").first.click()
            else:
                if cal_year < target_year or (cal_year == target_year and cal_month_idx < target_month_idx):
                    from_calendar.locator(".icon-arrow-right14").first.click()
                else:
                    from_calendar.locator(".icon-arrow-left12").first.click()
            target_page.wait_for_timeout(500)
            
        # Navigate To-Calendar
        print("Aligning To-Calendar month...")
        for attempt in range(24):
            header_text = to_calendar.evaluate("el => el.innerText").strip()
            lines = [l.strip() for l in header_text.split("\n") if "-" in l]
            if not lines:
                break
            cal_header = lines[0]
            if target_header in cal_header:
                break
            cal_year, cal_month_idx = parse_header(cal_header)
            if cal_year is None:
                to_calendar.locator(".icon-arrow-right14").first.click()
            else:
                if cal_year < target_year or (cal_year == target_year and cal_month_idx < target_month_idx):
                    to_calendar.locator(".icon-arrow-right14").first.click()
                else:
                    to_calendar.locator(".icon-arrow-left12").first.click()
            target_page.wait_for_timeout(500)
            
        # 5. Click Day 1 on From-Calendar
        print(f"Clicking {current_month_abbr} 1st on From-Calendar...")
        day_1_element = from_calendar.locator('td[class="day"]:has-text("1")')
        day_1_found = False
        for i in range(day_1_element.count()):
            if day_1_element.nth(i).inner_text().strip() == "1":
                day_1_element.nth(i).click()
                day_1_found = True
                break
        if not day_1_found:
            day_1_backup = from_calendar.locator('td:text-is("1")')
            if day_1_backup.count() > 0:
                day_1_backup.first.click()
                
        # 6. Click Today on To-Calendar
        print(f"Clicking {current_month_abbr} {current_day} on To-Calendar...")
        day_today_element = to_calendar.locator(f'td[class="day"]:has-text("{current_day}")')
        day_today_found = False
        for i in range(day_today_element.count()):
            if day_today_element.nth(i).inner_text().strip() == str(current_day):
                day_today_element.nth(i).click()
                day_today_found = True
                break
        if not day_today_found:
            day_today_backup = to_calendar.locator(f'td:text-is("{current_day}")')
            if day_today_backup.count() > 0:
                day_today_backup.first.click()
                
        target_page.wait_for_timeout(1000)
        
        # 7. Click Apply inside calendar
        print("Clicking 'Apply' button inside calendar...")
        apply_btn = target_page.locator(".apply-range-date")
        if apply_btn.count() > 0:
            apply_btn.first.click()
            print("Clicked Apply! Waiting 3 seconds for dashboard to begin query reload...")
            target_page.wait_for_timeout(3000)
            print("Custom date range successfully applied!")
        else:
            print("Warning: Could not find 'Apply' button inside calendar.")
            
        # 8. Open main Export dropdown
        print("\nOpening main 'Export' dropdown in header...")
        dropdown_opened = False
        
        # Check if Excel option is already visible to confirm dropdown is open
        excel_option = target_page.locator("text=Excel:visible").first
        if excel_option.count() > 0:
            print("Export dropdown is already open!")
            dropdown_opened = True
        else:
            print("Dropdown is not open. Searching for the main header 'Export' button...")
            main_export_btn = None
            for attempt in range(25):
                btn = target_page.locator("button.sc-iQZIRz.dJYOnr:has-text('Export'):visible").first
                if btn.count() == 0:
                    btn = target_page.locator("button:has-text('Export'):visible").first
                
                if btn.count() > 0:
                    main_export_btn = btn
                    print("Success: Found the main 'Export' button!")
                    break
                
                print(f"Export button not visible yet (Attempt {attempt+1}/25). Waiting...")
                target_page.wait_for_timeout(1000)
                
            if main_export_btn:
                print("Clicking header 'Export' button to open dropdown...")
                main_export_btn.click(force=True)
                print("Waiting for dropdown to render...")
                try:
                    target_page.wait_for_selector("text=Excel", state="visible", timeout=15000)
                    print("Export dropdown successfully opened!")
                    dropdown_opened = True
                except Exception:
                    print("Warning: Timed out waiting for the dropdown to render.")
            else:
                print("Error: Could not locate the main header 'Export' button.")
                return
            
        if dropdown_opened:
            print("\nSuccessfully opened the Export dropdown. Stopping automation here.")
        else:
            print("\nWarning: Could not open the Export dropdown menu. Stopping automation here.")

def main():
    if not is_chrome_running():
        launch_chrome()
    else:
        print("Chrome is already running on port 9222.")
        
    # Run the full automated sequence
    run_automation()

if __name__ == "__main__":
    main()
