import os
import json
import time
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

print("Starting script...")

# ----- GOOGLE SHEET SETUP -----
print("Starting Google Sheet setup...")
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Read the JSON key from GitHub Actions secret
key_json = os.environ["GOOGLE_APPLICATION_CREDENTIALS_JSON"]
creds_dict = json.loads(key_json)
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

sheet = client.open("activity_tracker").sheet1  # Replace with your sheet name
data = sheet.get_all_records()
print("Google Sheet setup done.")

# ----- SELENIUM SETUP -----
print("Setting up Selenium Chrome driver...")

chrome_options = Options()
chrome_options.add_argument("--headless=new")  # Headless mode
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

# Initialize Chrome driver
driver = webdriver.Chrome(options=chrome_options)

print("Chrome setup done. Opening WhatsApp Web...")

# Open WhatsApp Web in first tab
driver.get("https://web.whatsapp.com")
time.sleep(10)  # Wait for WhatsApp to load

# ----- MAIN LOOP -----
for j, row in enumerate(data, start=2):  # Assuming row 1 is header
    org_id = row["org id"]
    group_name = row["group_Name"]
    status = row["Status"]

    if status.lower() == "msg_sent":
        continue

    print(f"\nProcessing row {j} | Org ID: {org_id} | Group: {group_name}")

    # --- Open Retool ---
    driver.execute_script("window.open('');")  # Open new tab
    driver.switch_to.window(driver.window_handles[1])
    retool_url = f"https://getpp.retool.com/apps/4a16ac34-6f4a-11ef-a198-97d135a29ef7/Powerplay/ActivityReport/{org_id}"
    driver.get(retool_url)

    try:
        WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.XPATH, "//h1[text()='Activity Report']"))
        )
        print("Retool loaded.")
    except Exception as e:
        print("Timeout or error loading Retool:", e)

    time.sleep(5)

    # Search organization in Retool
    try:
        search_box = driver.find_element(By.XPATH, "//input[@placeholder='Search by Name or ID']")
        search_box.clear()
        search_box.send_keys(f"{org_id}" + Keys.ENTER)
        time.sleep(5)
        print("Organization searched in Retool.")
    except Exception as e:
        print("Error searching org in Retool:", e)

    # Scroll to table
    table_found = False
    for _ in range(10):
        try:
            table = driver.find_element(By.XPATH, "//div[@data-testid='RetoolGrid:table137']")
            driver.execute_script("arguments[0].scrollIntoView();", table)
            table_found = True
            print("Table found.")
            break
        except:
            driver.execute_script("window.scrollBy(0, 500);")
            time.sleep(1)

    if not table_found:
        print("Table not found, skipping this org.")
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        continue

    # Screenshot Retool page
    screenshot_path = f"/tmp/screenshot_{org_id}.png"
    driver.save_screenshot(screenshot_path)
    print(f"Screenshot saved: {screenshot_path}")

    # --- Send via WhatsApp ---
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(5)

    try:
        # Search for WhatsApp group
        search_box = driver.find_element(By.XPATH, "//div[@aria-label='Search input textbox']")
        search_box.click()
        search_box.send_keys(group_name)
        search_box.send_keys(Keys.ENTER)
        time.sleep(3)

        # Attach image
        attach_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//button[@title="Attach"]'))
        )
        attach_btn.click()
        time.sleep(1)

        image_input = driver.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_input.send_keys(screenshot_path)
        time.sleep(2)

        # Click send
        send_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@role="button" and @aria-label="Send"]'))
        )
        send_btn.click()
        time.sleep(2)
        print("Screenshot sent on WhatsApp.")

        # Update Google Sheet
        sheet.update_cell(j, 3, "msg_sent")  # Status column
        sheet.update_cell(j, 4, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))  # Sent timestamp
        print(f"Google Sheet updated for row {j}")

    except Exception as e:
        print(f"Error sending WhatsApp message for {group_name}: {e}")
        sheet.update_cell(j, 3, "error")

    # Close Retool tab
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

# Quit driver
driver.quit()
print("Script completed.")
