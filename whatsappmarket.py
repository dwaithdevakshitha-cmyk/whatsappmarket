import pandas as pd
import time
import sys
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import InvalidSessionIdException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

# --- USER CONFIGURATION ---
EXCEL_PATH = r"C:\Users\HP\Documents\Trip flux marketing\USA work Permit_1_to_40_06_03_2026.xlsx"
VIDEO_PATH = r"C:\Users\HP\Documents\Trip flux marketing\WhatsApp Video 2026-03-14 at 5.19.52 PM.mp4"
MESSAGE_TEXT = (
    "Finding a budget-friendly villa within a 40-minute area drive from Hyderabad's "
    "Financial District and HITEC City can be challenging, as these area are prime "
    "real estate zones with higher property prices. However, exploring nearby "
    "localities may offer more affordable options. Here are some suggestions:"
)
# --------------------------

# Load Excel - Searching for the header row containing 'Contact No.'
try:
    raw_df = pd.read_excel(EXCEL_PATH, header=None)
    header_row_index = 0
    for i, row in raw_df.iterrows():
        row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
        if "CONTACT NO." in row_str:
            header_row_index = i
            break
    
    df = pd.read_excel(EXCEL_PATH, header=header_row_index)
    df.columns = [str(c).upper().strip() for c in df.columns]
    print(f"Loaded Excel. Header found at row {header_row_index}. Columns: {df.columns.tolist()}")
except Exception as e:
    print(f"Error loading Excel: {e}")
    sys.exit(1)

if not os.path.exists(VIDEO_PATH):
    print(f"ERROR: Video file not found at {VIDEO_PATH}")
    sys.exit(1)

# Setup Chrome
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get("https://web.whatsapp.com")

print("Please scan the QR code to login...")
try:
    wait = WebDriverWait(driver, 300)
    wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="side"] | //div[@data-testid="chat-list"]')))
    print("Login successful! Starting to send videos and messages...")
except Exception as e:
    print(f"Login timed out or failed: {e}")
    driver.quit()
    sys.exit(1)

phone_column = 'CONTACT NO.'
if phone_column in df.columns:
    df = df.dropna(subset=[phone_column])
else:
    print(f"Error: Could not find '{phone_column}' column in Excel.")
    driver.quit()
    sys.exit(1)

for index, row in df.iterrows():
    try:
        raw_phone = str(row[phone_column]).split('.')[0].strip()
        phone = "".join(filter(str.isdigit, raw_phone))
        
        if not phone or len(phone) < 10:
            print(f"[{index+1}] Skipping invalid phone number: {raw_phone}")
            continue

        print(f"[{index+1}] Preparing to send to: {phone}")
        # Build URL without text - IMPORTANT: if we add text to URL, it might conflict with caption
        url = f"https://web.whatsapp.com/send?phone={phone}"
        driver.get(url)

        wait = WebDriverWait(driver, 80) # High timeout for slow chat loads
        
        # 1. Wait for page to load
        wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="main"]')))
        time.sleep(4)

        # 2. Upload video using the file input
        try:
            # Look for the media input specifically
            file_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')))
            file_input.send_keys(VIDEO_PATH)
            print(f"[{index+1}] Video file selected. Waiting for upload preview...")
        except Exception as e:
            print(f"[{index+1}] Could not find file input: {e}")
            continue

        # 3. Wait for the caption box to appear in the preview screen
        caption_box_paths = [
            '//div[@aria-placeholder="Add a caption"]',
            '//div[@contenteditable="true"][@data-tab="10"]',
            '//div[@data-testid="caption-keyboard-input"]',
            '//footer//div[@contenteditable="true"]'
        ]
        
        caption_box = None
        for path in caption_box_paths:
            try:
                caption_box = wait.until(EC.element_to_be_clickable((By.XPATH, path)))
                if caption_box: break
            except:
                continue
        
        if caption_box:
            time.sleep(2)
            # Use ActionChains for more reliable typing into contenteditable
            actions = ActionChains(driver)
            actions.move_to_element(caption_box).click().pause(1)
            actions.send_keys(MESSAGE_TEXT).pause(1)
            actions.send_keys(Keys.ENTER).perform()
            print(f"Successfully sent video and message to: {phone}")
        else:
            print(f"[{index+1}] Could not find caption box after upload. Forcing 'Enter'...")
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)

    except InvalidSessionIdException:
        print("Browser window was closed. Stopping.")
        break
    except Exception as e:
        print(f"Failed to send to {phone}: {e}")

    # Long delay for video processing and sending
    time.sleep(20)

print("All tasks finished.")
driver.quit()