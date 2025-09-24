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

driver = webdriver.Chrome(options=chrome_options)

print("Chrome driver setup done.")
print("Opening WhatsApp Web...")

driver.get("https://web.whatsapp.com")
time.sleep(10)

for j, row in enumerate(data, start=2):
    org_id = row["org id"]
    group_name = row["group_Name"]
    status = row["Status"]

    if status.lower() == "msg_sent":
        continue

    print(f"Processing row {j} - {group_name}")

    # --- RETOOL SCREENSHOT ---
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.get(f"https://getpp.retool.com/apps/4a16ac34-6f4a-11ef-a198-97d135a29ef7/Powerplay/ActivityReport/{org_id}")
    print("Try block executing of retool")
    try:
        print("Tried try block retool")
        WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.XPATH, "//h1[text()='Activity Report']"))
        )
    except Exception as e:
        print("Error loading Retool:", e)

    time.sleep(5)
    screenshot_path = f"/tmp/screenshot_{org_id}.png"
    driver.save_screenshot(screenshot_path)
    print(f"Screenshot saved: {screenshot_path}")

    # --- WHATSAPP ---
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(3)

    try:
        search_box = driver.find_element(By.XPATH, "//div[@aria-label='Search input textbox']")
        search_box.click()
        search_box.send_keys(group_name)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        attach_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//button[@title="Attach"]'))
        )
        attach_btn.click()
        time.sleep(1)

        image_input = driver.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_input.send_keys(screenshot_path)
        time.sleep(2)

        send_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@role="button" and @aria-label="Send"]'))
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

    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

driver.quit()
print("Script completed.")
