import time
import os
import pandas as pd
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from openpyxl import Workbook
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
import re
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Load input Excel (Note: This file must exist in the GitHub runner workspace)
df_input = pd.read_excel("MD_Case_Search_Date_Range_Scrapping_Tool_Input.xlsx")

# Generate dynamic output filename with Date and Time
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_filename = f"MD_Case_Search_Output_{current_time}.xlsx"

# Create workbook
wb = Workbook()
ws = wb.active
ws.title = "Results"

# Write headers
headers = ["S.No", "County", "Estate Number", "Filing Date", "Date of Death", "Type", "Status", "Name"]
ws.append(headers)

sno = 1

# Setup Selenium WebDriver for Headless GitHub Actions
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")           # CRITICAL for GitHub Actions
options.add_argument("--no-sandbox")             # CRITICAL for Linux Runners
options.add_argument("--disable-dev-shm-usage")  # Prevents memory crashes
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

# Pull Proxy credentials from environment variables (GitHub Secrets) 
# Provides a fallback for local testing
PROXY_USER = os.getenv("WEBSHARE_USER", "yodzgxmr")
PROXY_PASS = os.getenv("WEBSHARE_PASS", "6h1gsqrqowmc")
PROXY_IP = os.getenv("WEBSHARE_IP", "23.95.150.145")
PROXY_PORT = os.getenv("WEBSHARE_PORT", "6114")

if not PROXY_USER:
    print("[!] No Proxy credentials found! Continuing without proxies...")
    proxy_options = {}
else:
    proxy_options = {
        'proxy': {
            'http': f'http://{PROXY_USER}:{PROXY_PASS}@{PROXY_IP}:{PROXY_PORT}',
            'https': f'https://{PROXY_USER}:{PROXY_PASS}@{PROXY_IP}:{PROXY_PORT}',
            'no_proxy': 'localhost,127.0.0.1'
        }
    }

print("Starting browser with Proxy in headless mode...")

# Launch with Proxy Configuration
driver = webdriver.Chrome(
    options=options,
    seleniumwire_options=proxy_options if proxy_options else None
)

def safe_text(td_element):
    text = td_element.text.strip()
    return text if text else ""

def go_to_page(driver, target_page):
    success = False
    tries = 0
    while not success and tries < 5:
        tries += 1
        try:
            pager_row = driver.find_element(By.XPATH, '//tr[@class="grid-pager"]/td')
            elements = pager_row.find_elements(By.XPATH, '*')

            for elem in elements:
                tag = elem.tag_name
                text = elem.text.strip()

                if tag == 'span' and text == str(target_page):
                    return True

                elif tag == 'a' and text == str(target_page):
                    driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                    ActionChains(driver).move_to_element(elem).click().perform()
                    time.sleep(2)
                    return True

            for elem in reversed(elements):
                if elem.tag_name == 'a' and elem.text.strip() == "...":
                    driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                    ActionChains(driver).move_to_element(elem).click().perform()
                    time.sleep(2)
                    break
        except:
            print(f"[!] Error navigating to page {target_page}")
    return False

try:
    for index, row in df_input.iterrows():
        county = str(row["County"]).strip()
        date_from = pd.to_datetime(row["From"]).strftime("%m/%d/%Y")
        date_to = pd.to_datetime(row["To"]).strftime("%m/%d/%Y")

        print(f"Loading Maryland Case Search for County: {county}")
        driver.get("https://registers.maryland.gov/RowNetWeb/Estates/frmEstateSearch2.aspx")
        time.sleep(2)

        try:
            county_dropdown = Select(driver.find_element(By.XPATH, '//*[@id="cboCountyId"]'))
            matched = False
            for option in county_dropdown.options:
                if county.lower() in option.text.lower():
                    county_dropdown.select_by_visible_text(option.text)
                    matched = True
                    break
            if not matched:
                print(f"[!] County '{county}' not found.")
                continue
        except:
            print(f"[!] Failed to select county")
            continue

        driver.find_element(By.XPATH, '//*[@id="DateOfFilingFrom"]').clear()
        driver.find_element(By.XPATH, '//*[@id="DateOfFilingFrom"]').send_keys(date_from)
        driver.find_element(By.XPATH, '//*[@id="DateOfFilingTo"]').clear()
        driver.find_element(By.XPATH, '//*[@id="DateOfFilingTo"]').send_keys(date_to)

        driver.find_element(By.XPATH, '//*[@id="cmdSearch"]').click()
        time.sleep(2)

        try:
            status_text = driver.find_element(By.XPATH, '//*[@id="tblStatus"]/tbody/tr/td').text
            ws.append([status_text]) 
            
            match = re.search(r'Page \d+ of (\d+).*?\((\d+) RECORDS', status_text)
            if match:
                total_pages = int(match.group(1))
                total_records_reported = int(match.group(2))
                print(f"[{county}] Total pages: {total_pages}, Total records (reported): {total_records_reported}")
                records_scraped = 0
            else:
                print(f"[!] Unexpected status format: '{status_text}'")
                continue
        except:
            print(f"[!] No results found for county")
            continue

        current_page = 1
        while current_page <= total_pages:
            print(f"[{county}] Scraping page {current_page}")

            rows = driver.find_elements(By.XPATH, '//*[@id="dgSearchResults"]/tbody/tr[position()>1 and position()<last()]')
            for i, row_elem in enumerate(rows, start=2):
                try:
                    tds = row_elem.find_elements(By.TAG_NAME, 'td')
                    if len(tds) == 7:
                        record = [
                            str(sno),
                            safe_text(tds[0]),
                            safe_text(tds[1]),
                            safe_text(tds[2]),
                            safe_text(tds[3]),
                            safe_text(tds[4]),
                            safe_text(tds[5]),
                            safe_text(tds[6]),
                        ]
                        ws.append(record)
                        sno += 1
                        records_scraped += 1
                        time.sleep(0.3)
                except:
                    print(f"Error processing row {i}")

            wb.save(output_filename)
            current_page += 1
            if current_page <= total_pages:
                if not go_to_page(driver, current_page):
                    print(f"[!] Failed to navigate to page {current_page}")
                    break
        
        if records_scraped == total_records_reported:
            print(f"[{county}] ✅ Record count matches: {records_scraped} records scraped.")
        else:
            print(f"[{county}] ❌ Record count mismatch! Reported: {total_records_reported}, Scraped: {records_scraped}")

except Exception as e:
    print(f"[!] Unexpected error: {e}")

wb.save(output_filename)
driver.quit()
print(f"✅ Automated Script Run Completed! Data saved to '{output_filename}'")

# ==========================================
# 📧 EMAIL AUTOMATION
# ==========================================
def send_email_report(filepath):
    # GitHub Secrets configuration (uses your email as default)
    SENDER_EMAIL = os.getenv("SENDER_EMAIL", "samp93178@gmail.com") 
    SENDER_APP_PASSWORD = os.getenv("SENDER_APP_PASSWORD", "") # Your 16-digit Google App Password
    RECIPIENT_EMAIL = os.getenv("RECIPIENT_EMAIL", "psantosh1999@gmail.com")
    
    if not SENDER_APP_PASSWORD:
        print("⚠️ No Email App Password found in environment setting. Skipping email dispatch.")
        return

    print("📤 Sending report via Email...")
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECIPIENT_EMAIL
        msg['Subject'] = f"✅ Scraper Report: Maryland Case Search ({current_time})"
        
        body = "The daily data scraping task has completed automatically. Please find the attached report."
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach the Excel file
        with open(filepath, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {filepath}")
            msg.attach(part)
        
        # Connect to Gmail SMTP Server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
        server.send_message(msg)
        server.quit()
        print(f"✅ Email securely sent to {RECIPIENT_EMAIL}!")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")

# Call the email function at the very end
send_email_report(output_filename)
