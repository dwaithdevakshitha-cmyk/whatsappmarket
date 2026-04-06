import typing

import pandas as pd  # type: ignore
import time
import sys
import os
from selenium import webdriver  # type: ignore
from selenium.webdriver.chrome.service import Service  # type: ignore
from selenium.webdriver.chrome.options import Options  # type: ignore
from webdriver_manager.chrome import ChromeDriverManager  # type: ignore
from selenium.webdriver.common.by import By  # type: ignore
from selenium.webdriver.common.keys import Keys  # type: ignore
from selenium.webdriver.support.wait import WebDriverWait  # type: ignore
from selenium.webdriver.support import expected_conditions as EC  # type: ignore
from selenium.common.exceptions import InvalidSessionIdException, TimeoutException  # type: ignore
from selenium.webdriver.common.action_chains import ActionChains  # type: ignore

# --- USER CONFIGURATION ---
EXCEL_PATH = r"C:\Users\HP\Downloads\associate 7.xlsx"
MEDIA_PATHS = [
    r"C:\Users\HP\Documents\Trip flux marketing banners\gokarna.png",
    r"C:\Users\HP\Documents\Trip flux marketing banners\kashi.png",
    r"C:\Users\HP\Documents\Trip flux marketing banners\kashi-nepal-mukthinath yathra.png",
    r"C:\Users\HP\Documents\Trip flux marketing banners\manali.png",
    r"C:\Users\HP\Documents\Trip flux marketing banners\Mansarovar package.png",
    r"C:\Users\HP\Documents\Trip flux marketing banners\Shri Ramayana yatra Sri Lanka Banner.png",
    r"C:\Users\HP\Documents\Trip flux marketing banners\Yamuna pushkaralu.png"
]
MESSAGE_TEXT = """Job SummaryThe Associate Travel Sales Executive supports the sales team by assisting customers with travel-related products and services. The role focuses on understanding client needs, promoting travel packages, closing sales, and ensuring excellent customer service before and after booking.Roles and ResponsibilitiesAssist customers with inquiries related to tour packages, flights, hotels, visas, and travel insuranceUnderstand customer preferences and recommend suitable travel productsExplain itineraries, pricing, inclusions, and exclusions clearly to clientsGenerate and follow up on sales leads via phone, email, or in personAchieve assigned sales targets and contribute to team revenue goalsMaintain accurate customer records and sales reportsHandle customer complaints or issues professionally and escalate when neededStay updated on travel trends, destinations, and promotional offersRequired Skills and QualificationsStrong communication and interpersonal skillsBasic knowledge of travel destinations and booking systems (preferred)Sales-oriented mindset with customer-focused approachAbility to work under targets and deadlines10th Pass , Intermediate, Graduate or diploma in Travel, Tourism, Sales, or related field (preferred)"""
# --------------------------

# Load Excel
try:
    # First, detect the header row
    raw_df = pd.read_excel(EXCEL_PATH, header=None)
    header_row_index = 0
    for i, row in raw_df.iterrows():
        row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
        if any(kw in row_str for kw in ["MOBILE NUMBER", "PHONE NUMBER", "CONTACT"]):
            header_row_index = int(str(i))
            break
    
    # Load the actual data using the detected header
    df = pd.read_excel(EXCEL_PATH, header=header_row_index)
    
    # Standardize column names (remove extra spaces)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Add/Check Delivery Status column
    STATUS_COLUMN = 'DELIVERY STATUS'
    if STATUS_COLUMN not in df.columns:
        df[STATUS_COLUMN] = "PENDING"
    else:
        df[STATUS_COLUMN] = df[STATUS_COLUMN].astype(str)
    
    print(f"Loaded Excel. Header found at row {header_row_index}. Total records: {len(df)}")
except Exception as e:
    print(f"Error loading Excel: {e}")
    sys.exit(1)

for p in MEDIA_PATHS:
    if not os.path.exists(p):
        print(f"ERROR: Media file not found at {p}")
        sys.exit(1)

# Setup Chrome
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver: typing.Any = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
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

# Detect phone column
phone_column = None
for col in df.columns:
    if any(kw in col.upper() for kw in ["MOBILE NUMBER", "PHONE NUMBER", "CONTACT"]):
        phone_column = col
        break

if not phone_column:
    print(f"Error: Could not find any phone number column in Excel. Available columns: {df.columns.tolist()}")
    driver.quit()
    sys.exit(1)

# --- CLEANING AND FILTERING ---
# 1. Clean phone numbers first to ensure accurate deduplication
def _clean_phone(x):
    p = str(x).split(',')[0].split('/')[0].strip()
    p = p.split('.')[0].strip()
    return "".join(filter(str.isdigit, p))

df[phone_column] = df[phone_column].apply(_clean_phone)

# 1.5. Check global master list to prevent cross-file repeat sends
MASTER_LOG_PATH = "sent_numbers_history.txt"
if os.path.exists(MASTER_LOG_PATH):
    with open(MASTER_LOG_PATH, "r") as f:
        global_handled = set(line.strip() for line in f if line.strip())
    df.loc[df[phone_column].isin(global_handled) & (df[STATUS_COLUMN].isin(['PENDING', 'nan']) | df[STATUS_COLUMN].isna()), STATUS_COLUMN] = "ALREADY_PROCESSED"

# 1.6. Propagate 'SENT' to all duplicates inside this specific file
handled_statuses = ['SENT', 'ALREADY_PROCESSED', 'INVALID_NUMBER', 'FAILED_IMAGE', 'INVALID_FORMAT']
handled_numbers = df[df[STATUS_COLUMN].isin(handled_statuses)][phone_column].unique()
df.loc[df[phone_column].isin(handled_numbers) & (df[STATUS_COLUMN].isin(['PENDING', 'nan']) | df[STATUS_COLUMN].isna()), STATUS_COLUMN] = "SENT_DUPLICATE"

# 2. Remove duplicates based on phone number
initial_count = len(df)
df = df.drop_duplicates(subset=[phone_column], keep='first')
if len(df) < initial_count:
    print(f"Removed {initial_count - len(df)} duplicate phone numbers.")

# 2. Skip already sent or invalid numbers
retry_statuses = ['PENDING', 'FAILED', 'ERROR', 'FAILED_IMAGE', 'TIMEOUT', 'nan']
df_to_process = df[df[STATUS_COLUMN].isin(retry_statuses) | df[STATUS_COLUMN].isna()]
print(f"Skipping {len(df) - len(df_to_process)} numbers already marked as SENT or INVALID.")

def save_excel_progress():
    """Saves the current dataframe back to the Excel file."""
    try:
        df.to_excel(EXCEL_PATH, index=False)
    except Exception as e:
        print(f"\nWARNING: Could not save to Excel: {e}")
        print("Please CLOSE the Excel file if it is open so the script can update status!")

for i_idx, (index, row) in enumerate(df_to_process.iterrows()):
    phone = ""
    try:
        phone = str(row[phone_column])
        
        if not phone or len(phone) < 10:
            print(f"[{i_idx+1}] Skipping invalid phone number format: {phone}")
            df.at[index, STATUS_COLUMN] = "INVALID_FORMAT"
            save_excel_progress()
            continue

        print(f"[{i_idx+1}] Preparing to send to: {phone}")
        url = f"https://web.whatsapp.com/send?phone={phone}"
        driver.get(url)  # type: ignore

        wait = WebDriverWait(driver, 80)
        
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="main"] | //div[@role="dialog"] | //div[@data-animate-modal-popup="true"] | //button[@data-testid="popup-controls-ok"]')))
            time.sleep(2)
            
            is_invalid = False
            # Check for any dialog or popup indicating an invalid number
            popups = driver.find_elements(By.XPATH, '//div[@role="dialog"] | //div[@data-animate-modal-popup="true"] | //div[contains(@class, "overlay")]')
            for popup in popups:
                popup_text = popup.text.lower()
                error_keywords = ["invalid", "url", "incorrect", "not registered", "not on whatsapp", "whatsapp account", "isn't on whatsapp"]
                if any(kw in popup_text for kw in error_keywords):
                    is_invalid = True
                    print(f"[{i_idx+1}] Detected invalid number alert: {popup_text}")
                    # Try specifically to find an "OK" or "Close" button
                    try:
                        ok_btn = popup.find_elements(By.XPATH, './/button | .//div[@role="button"] | .//div[@data-testid="popup-controls-ok"]')
                        if ok_btn:
                            driver.execute_script("arguments[0].click();", ok_btn[0])
                            time.sleep(1)
                    except:
                        pass
                    break
            
            # Additional fallback check for the specific 'OK' button that sometimes appears alone
            if not is_invalid:
                fallback_btns = driver.find_elements(By.XPATH, '//button[@data-testid="popup-controls-ok"] | //div[@data-testid="popup-controls-ok"] | //button[normalize-space()="OK"]')
                if fallback_btns:
                    is_invalid = True
                    print(f"[{i_idx+1}] Detected standalone 'OK' button for invalid status.")
                    try:
                        driver.execute_script("arguments[0].click();", fallback_btns[0])
                        time.sleep(1)
                    except:
                        pass

            if is_invalid:
                print(f"[{i_idx+1}] Skipping: Number {phone} does not have a WhatsApp account.")
                df.at[index, STATUS_COLUMN] = "INVALID_NUMBER"
                with open(MASTER_LOG_PATH, "a") as f:
                    f.write(phone + "\n")
                save_excel_progress()
                continue
                
        except TimeoutException:
            print(f"[{i_idx+1}] Timeout waiting for {phone}. Skipping.")
            df.at[index, STATUS_COLUMN] = "TIMEOUT"
            save_excel_progress()
            continue

        # 2. Send text separately (chat box supports 65k characters, skipping image caption limit of 1024 entirely)
        try:
            chat_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@title="Type a message"] | //div[@contenteditable="true"][@data-tab="10"] | //footer//div[@contenteditable="true"]')))
            actions = ActionChains(driver)
            actions.move_to_element(chat_box).click().pause(1)
            for part in MESSAGE_TEXT.split('\n'):
                actions.send_keys(part)
                actions.key_down(Keys.SHIFT).send_keys(Keys.ENTER).key_up(Keys.SHIFT)
            actions.send_keys(Keys.ENTER).perform()
            time.sleep(2)
        except Exception as e:
            print(f"[{i_idx+1}] Could not send text: {e}")
            df.at[index, STATUS_COLUMN] = "FAILED_TEXT"
            save_excel_progress()
            continue

        # 3. Send images directly afterwards
        try:
            import subprocess
            if MEDIA_PATHS:
                abs_paths = [os.path.abspath(p) for p in MEDIA_PATHS if os.path.exists(p)]
                if abs_paths:
                    # Explicitly use the PowerShell clipboard array method for the images
                    ps_paths = ", ".join([f"'{p.replace('\'', '\'\'')}'" for p in abs_paths])
                    subprocess.run(['powershell', '-command', f"Set-Clipboard -Path {ps_paths}"], shell=True)
                    
                    chat_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@title="Type a message"] | //div[@contenteditable="true"][@data-tab="10"] | //footer//div[@contenteditable="true"]')))
                    actions = ActionChains(driver)
                    actions.move_to_element(chat_box).click().pause(1).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                    
                    # Wait specifically for the image modal to fully open
                    time.sleep(8)
                    
                    # First attempt: simply press ENTER which natively sends the images from the open modal
                    try:
                        actions = ActionChains(driver)
                        actions.send_keys(Keys.ENTER).perform()
                    except Exception as e:
                        print(f"[{i_idx+1}] Enter key failed: {e}")
                    
                    time.sleep(3)
                    
                    # Second attempt (Fallback): visually locate and click the Send button
                    try:
                        send_btn = driver.find_element(By.XPATH, '//div[@aria-label="Send" or @aria-label="Send " or @data-testid="send"] | //span[@data-icon="send"]/..')
                        actions = ActionChains(driver)
                        actions.move_to_element(send_btn).click().perform()
                    except Exception:
                        pass
                    time.sleep(5)
        except Exception as e:
            print(f"[{i_idx+1}] Could not upload images: {e}")
            df.at[index, STATUS_COLUMN] = "FAILED_IMAGE"
            save_excel_progress()
            continue
        
        print(f"[{i_idx+1}] Successfully sent to: {phone}")
        df.at[index, STATUS_COLUMN] = "SENT"
        with open(MASTER_LOG_PATH, "a") as f:
            f.write(phone + "\n")
        save_excel_progress()

    except InvalidSessionIdException:
        print("Browser window was closed. Stopping.")
        break
    except Exception as e:
        print(f"Failed to send to {phone}: {e}")
        df.at[index, STATUS_COLUMN] = "ERROR"
        save_excel_progress()

    time.sleep(20)

print("\nAll tasks finished.")
driver.quit()  # type: ignore