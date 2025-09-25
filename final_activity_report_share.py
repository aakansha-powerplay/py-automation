import time
import datetime
import gspread
import os
import json
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

print("🚀 Starting script...")

# ----- GOOGLE SHEET SETUP -----
print("📑 Setting up Google Sheet...")
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_json = os.environ["GSPREAD_KEY"]
creds_dict = json.loads(creds_json)

creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

sheet = client.open("activity_tracker").sheet1  # Replace with your sheet name
data = sheet.get_all_records()
print(f"✅ Loaded {len(data)} rows from Google Sheet.")

# ----- SELENIUM CHROME SETUP -----
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=chrome_options)
print("✅ Chrome driver setup done.")

# ----- WHATSAPP LOGIN -----
print("🌐 Opening WhatsApp Web...")
driver.get("https://web.whatsapp.com")

# Load cookies if present
whatsapp_cookies = os.environ.get("WHATSAPP_COOKIES")
if whatsapp_cookies:
    try:
        cookies = json.loads(whatsapp_cookies)
        driver.delete_all_cookies()
        for cookie in cookies:
            driver.add_cookie(cookie)
        print(f"✅ Loaded {len(cookies)} WhatsApp cookies.")
        driver.refresh()
    except Exception as e:
        print(f"⚠️ Error loading cookies: {e}")

# Check if QR or Chat loaded
try:
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//div[@role='textbox']"))
    )
    print("✅ WhatsApp session restored successfully.")
except:
    try:
        qr_present = driver.find_element(By.XPATH, "//canvas[@aria-label='Scan me!']")
        print("⚠️ QR code detected — cookies are invalid/expired. Please login locally and update WHATSAPP_COOKIES secret.")
        driver.save_screenshot("/tmp/whatsapp_qr_detected.png")
    except:
        print("❌ WhatsApp neither loaded chat nor QR. Check network/cookies.")
        driver.save_screenshot("/tmp/whatsapp_login_failed.png")

# ----- PROCESS ROWS -----
for j, row in enumerate(data, start=2):
    org_id = row.get("org id")
    chat_name = row.get("group_Name")
    status = row.get("Status")

    if not org_id or not chat_name:
        print(f"⚠️ Row {j} skipped (missing org_id or chat_name).")
        continue

    if status and status.lower() == "msg_sent":
        continue

    print(f"➡️ Processing row {j} - Chat: {chat_name}")

    # --- RETOOL SCREENSHOT (PUBLIC LINK) ---
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    public_url = f"https://getpp.retool.com/embedded/public/0ff1d82d-d7b9-412e-9c43-2b0123a5529f/activity_report?org={org_id}"
    driver.get(public_url)

    screenshot_path = f"/tmp/screenshot_{org_id}_{int(time.time())}.png"
    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
    except Exception as e:
        print(f"⚠️ Retool public link failed for {org_id}: {e}")
        driver.save_screenshot(f"/tmp/retool_error_{org_id}.png")

    time.sleep(3)
    driver.save_screenshot(screenshot_path)
    print(f"📸 Screenshot saved: {screenshot_path}")

    # --- WHATSAPP MESSAGE ---
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(3)

    try:
        search_box = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='3']"))
        )
        search_box.click()
        search_box.clear()
        search_box.send_keys(chat_name)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        attach_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="clip"]'))
        )
        attach_btn.click()
        time.sleep(1)

        image_input = driver.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_input.send_keys(screenshot_path)
        time.sleep(2)

        send_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
        )
        send_btn.click()
        time.sleep(2)

        print(f"✅ Message sent to {chat_name}")
        sheet.update_cell(j, 3, "msg_sent")
        sheet.update_cell(j, 4, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    except Exception as e:
        print(f"❌ Error sending to {chat_name}: {e}")
        error_screenshot = f"/tmp/error_{org_id}_{int(time.time())}.png"
        driver.save_screenshot(error_screenshot)
        print(f"📸 Error screenshot saved: {error_screenshot}")
        sheet.update_cell(j, 3, "error")

    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

driver.quit()
print("🎯 Script completed.")
