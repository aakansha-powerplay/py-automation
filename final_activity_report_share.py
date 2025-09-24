import time
import datetime
import gspread
import os
import json
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

print("Starting script...")
print("Starting Google Sheet setup...")

# ----- GOOGLE SHEET SETUP -----
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load JSON directly from secret
creds_json = os.environ["GSPREAD_KEY"]
creds_dict = json.loads(creds_json)
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

sheet = client.open("activity_tracker").sheet1  # Replace with your sheet name
data = sheet.get_all_records()
print("Google Sheet setup done.")

# ----- SELENIUM CHROME SETUP -----
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=chrome_options)
print("Chrome driver setup done.")
print("Opening WhatsApp Web...")

driver.get("https://web.whatsapp.com")
print("Please scan QR code if needed (only once if not cached).")
time.sleep(10)  # Adjust if QR scan takes longer

# ----- MAIN LOOP -----
for j, row in enumerate(data, start=2):
    org_id = row["org id"]
    group_name = row["group_Name"]
    status = row["Status"]

    if status.lower() == "msg_sent":
        continue

    print(f"Processing row {j} - {group_name}")

    # --- RETOOL SCREENSHOT ---
    try:
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[1])
        driver.get(f"https://getpp.retool.com/apps/4a16ac34-6f4a-11ef-a198-97d135a29ef7/Powerplay/ActivityReport/{org_id}")

        print("Waiting for Retool to load...")
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        time.sleep(5)

        screenshot_path = f"/tmp/screenshot_{org_id}.png"
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot saved: {screenshot_path}")

    except Exception as e:
        print(f"Error loading Retool: {e}")
        screenshot_path = None

    # --- WHATSAPP ---
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(3)

    try:
        # Wait and find search box
        search_box = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@contenteditable='true' and @data-tab]")
            )
        )
        search_box.click()
        search_box.clear()
        search_box.send_keys(group_name)
        search_box.send_keys(Keys.ENTER)
        time.sleep(3)

        if screenshot_path:
            attach_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="clip"]'))
            )
            attach_btn.click()
            time.sleep(1)

            image_input = driver.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            image_input.send_keys(screenshot_path)
            time.sleep(2)

            send_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            )
            send_btn.click()
            time.sleep(2)

        print(f"Message sent to {group_name}")

        # Update sheet
        sheet.update_cell(j, 3, "msg_sent")
        sheet.update_cell(j, 4, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    except Exception as e:
        print(f"Error sending to {group_name}: {e}")
        sheet.update_cell(j, 3, "error")
        # Save screenshot on failure
        driver.save_screenshot(f"/tmp/error_{org_id}.png")

    # Close Retool tab
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

driver.quit()
print("Script completed.")
