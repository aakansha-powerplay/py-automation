import os
import json
import time
import traceback
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =========================
#  GOOGLE SHEETS SETUP
# =========================
print("🚀 Starting script...")
print("📑 Setting up Google Sheet...")
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_json = json.loads(os.getenv("GSPREAD_KEY"))
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
client = gspread.authorize(creds)

sheet = client.open("Activity Report Share").sheet1
data = sheet.get_all_records()
print(f"✅ Loaded {len(data)} rows from Google Sheet.")

# =========================
#  CHROME + SELENIUM SETUP
# =========================
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=chrome_options)
print("✅ Chrome driver setup done.")

# =========================
#  WHATSAPP LOGIN VIA COOKIES
# =========================
print("🌐 Opening WhatsApp Web...")
driver.get("https://web.whatsapp.com")
cookies_env = os.getenv("WHATSAPP_COOKIES")

if not cookies_env:
    print("❌ No cookies found in env. QR login required.")
else:
    try:
        cookies = json.loads(cookies_env)
        for cookie in cookies:
            driver.add_cookie(cookie)
        print(f"✅ Loaded {len(cookies)} WhatsApp cookies. Reloading page...")
        driver.refresh()
        time.sleep(5)
        # Debug check: QR or chat screen?
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "canvas"))
            )
            print("⚠️ WhatsApp shows QR code — cookies may be invalid or expired.")
        except:
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='grid']"))
                )
                print("✅ WhatsApp chat list loaded successfully.")
            except:
                print("❌ WhatsApp neither loaded chat nor QR. Check network/cookies.")
    except Exception as e:
        print(f"❌ Failed to load cookies: {e}")

# =========================
#  PROCESS EACH ROW
# =========================
for j, row in enumerate(data, start=2):
    try:
        org_id = row.get("org id", "").strip()
        group_name = row.get("group_Name", "").strip()
        status = row.get("Status", "").strip()

        print(f"➡️ Processing row {j} - Chat: {group_name} | Org: {org_id}")

        # Skip if org_id missing
        if not org_id:
            print(f"⚠️ Skipping row {j}: Missing org_id")
            continue

        # =========================
        #  OPEN RETOOL PUBLIC LINK
        # =========================
        public_url = f"https://getpp.retool.com/embedded/public/0ff1d82d-d7b9-412e-9c43-2b0123a5529f/activity_report?org={org_id}"
        driver.get(public_url)

        screenshot_path = f"/tmp/screenshot_{org_id}_{int(time.time())}.png"
        try:
            # Wait until Retool loads a widget container
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'retool-widget-container')]"))
            )
            print(f"✅ Retool dashboard loaded for {org_id}")
        except Exception as e:
            print(f"⚠️ Retool public link failed or slow for {org_id}: {e}")
            driver.save_screenshot(f"/tmp/retool_error_{org_id}.png")

        time.sleep(5)  # give UI a little more time
        driver.save_screenshot(screenshot_path)
        print(f"📸 Screenshot saved: {screenshot_path}")

        # =========================
        #  SEND VIA WHATSAPP
        # =========================
        try:
            driver.get("https://web.whatsapp.com")
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='grid']"))
            )
            search_box = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )
            search_box.clear()
            search_box.send_keys(group_name)
            time.sleep(3)

            # click first chat result
            chat = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[@title='" + group_name + "']"))
            )
            chat.click()

            # attach and send screenshot
            attach = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='clip']"))
            )
            attach.click()

            image_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
            image_input.send_keys(screenshot_path)
            time.sleep(3)

            send_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send']"))
            )
            send_button.click()

            print(f"✅ Sent screenshot to {group_name}")

        except Exception as e:
            print(f"❌ Error sending to {group_name}: {e}")
            driver.save_screenshot(f"/tmp/error_{org_id}_{int(time.time())}.png")

    except Exception as e:
        print(f"❌ Unexpected error in row {j}: {traceback.format_exc()}")

print("🎯 Script completed.")
driver.quit()
