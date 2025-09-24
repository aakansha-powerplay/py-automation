import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import ImageGrab
import datetime
import gspread
from selenium.webdriver.common.keys import Keys
from oauth2client.service_account import ServiceAccountCredentials

print("starting now")

# ----- CONFIGURATION -----
CHROMEDRIVER_PATH = "/usr/local/bin/chromedriver" 
CHROME_PROFILE_PATH = "/Users/aakansha/Library/Application Support/Google/Chrome/Default"  # adjust if different
PROFILE_DIR = "Default"  # or "Profile 1", etc.

#RETOOL_URL_TEMPLATE = "https://your-retool-url.com/org?orgId={}"  # 👈 Customize this

print("starting google sheet now")
# ----- GOOGLE SHEET SETUP -----
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("/Users/aakansha/credentials/google_credentials_service.json", scope)
client = gspread.authorize(creds)

sheet = client.open("activity_tracker").sheet1  # 👈 Replace with your sheet name
data = sheet.get_all_records()

print("done: google sheet ")

# ----- SELENIUM SETUP -----
# Set Chrome options to use your profile
options = Options()
options.add_argument(f"user-data-dir={CHROME_PROFILE_PATH}")
options.add_argument(f"profile-directory={PROFILE_DIR}")

# Optional: run headless
# options.add_argument("--headless=new")

# Initialize driver
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)


print("opening wa now")

#-------- open WA and retool ---------
# 1. Open WhatsApp Web in first tab
driver.get("https://web.whatsapp.com")
time.sleep(10)  # Wait for WA to load if needed

print("opening another window")


print("starting loop now ")

for j, row in enumerate(data, start=2):  # Assuming row 1 is header
    org_id = row["org id"]
    group_name = row["group_Name"]
    status = row["Status"]

    if status.lower() == "msg_sent":
        continue
    
    print(j)
    
    print("retool opening")
    
    print(f"https://getpp.retool.com/apps/4a16ac34-6f4a-11ef-a198-97d135a29ef7/Powerplay/ActivityReport/{org_id}")
    
    # Step 1: Open Retool
    # 2. Open new tab and go to another URL
    driver.execute_script("window.open('');")  # Open a new blank tab
    driver.switch_to.window(driver.window_handles[1])  # Switch to the new tab
    driver.get(f"https://getpp.retool.com/apps/4a16ac34-6f4a-11ef-a198-97d135a29ef7/Powerplay/ActivityReport/{org_id}")



    try:
        print("Waiting for Retool to load...")
        WebDriverWait(driver, 60).until(
            EC.any_of(
                EC.visibility_of_element_located((By.XPATH, "//h1[text()='Activity Report']"))  # chat area
            )
        )
    # %%

        print("Retool loaded.")

    except Exception as e:
        print("Timeout or error:", e)



    time.sleep(10)
    print("search org")

    search_box = driver.find_element(By.XPATH, "//input[@placeholder='Search by Name or ID']")
    search_box.clear()  # Optional: Clears any pre-existing text
    search_box.send_keys(f"{org_id}" + Keys.ENTER) 
    print("searching for org")
    time.sleep(10)

    print("scroll starting")

    # Scroll down until the table is visible
    table_found = False

    for i in range(10):  # Try scrolling up to 10 times
        try:
            table = driver.find_element(By.XPATH, "//div[@data-testid='RetoolGrid:table137']")
            driver.execute_script("arguments[0].scrollIntoView();", table)
            print("✅ Table found.")
            table_found = True
            break
        except:
            driver.execute_script("window.scrollBy(0, 500);")
            time.sleep(1)

    if not table_found:
        print("❌ Table not found.")
        driver.quit()
        
  # --- 2. Screenshot ---

    # Save screenshot of only the webpage area
    screenshot_path = f"/tmp/screenshot_{org_id}.png"
    driver.save_screenshot(screenshot_path)
    print(f"✅ Page screenshot saved to: {screenshot_path}")

  
   
   

    # --- 3. Go to WhatsApp ---
    driver.switch_to.window(driver.window_handles[0]) 
    time.sleep(8)

    try:
        print("✅ search contact!")
        
        search_xpath = "//div[@aria-label='Search input textbox']"
        search_box = driver.find_element(By.XPATH, search_xpath)
        search_box.click()
        search_box.send_keys(group_name)
        search_box.send_keys(Keys.ENTER)
        time.sleep(3)

        # Click the attach button
        wait_for = WebDriverWait(driver, 20)
        attach_btn = wait_for.until(EC.presence_of_element_located((By.XPATH, '//button[@title="Attach"]')))
        attach_btn.click()
        time.sleep(1)
        

        # Upload screenshot
        image_input = driver.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_input.send_keys(screenshot_path)
        time.sleep(2)

        # Click send button
        # Click the send button
        send_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@role="button" and @aria-label="Send"]'))
        )
        send_btn.click()
        time.sleep(2)
       
        print("✅ Screenshot sent!")
        #-------
        
        print(j)

        # --- 4. Update Google Sheet ---
        sheet.update_cell(j, 3, "msg_sent")  # Status column (C)
        sheet.update_cell(j, 4, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))  # Sent Date (D)

        print(f"✅ Message sent to {group_name}")

    except Exception as e:
        print(f"❌ Error with {group_name}: {e}")
        sheet.update_cell(i, 3, "error")


    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

driver.quit()