import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from openpyxl import Workbook
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import messagebox
import re

# Load input Excel (skip first row)
df_input = pd.read_excel("MD_Case_Search_Date_Range_Scrapping_Tool_Input.xlsx")

# Create workbook
wb = Workbook()
ws = wb.active
ws.title = "Results"

# Write headers
headers = ["S.No", "County", "Estate Number", "Filing Date", "Date of Death", "Type", "Status", "Name"]
ws.append(headers)

sno = 1

# Setup Selenium WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

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
                    return True  # Already on this page

                elif tag == 'a' and text == str(target_page):
                    driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                    ActionChains(driver).move_to_element(elem).click().perform()
                    time.sleep(2)
                    return True

            # Click last ellipsis if target page not in view
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

        driver.get("https://registers.maryland.gov/RowNetWeb/Estates/frmEstateSearch2.aspx")
        time.sleep(2)

        # Select County from dropdown
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

        # Set Date fields
        driver.find_element(By.XPATH, '//*[@id="DateOfFilingFrom"]').clear()
        driver.find_element(By.XPATH, '//*[@id="DateOfFilingFrom"]').send_keys(date_from)

        driver.find_element(By.XPATH, '//*[@id="DateOfFilingTo"]').clear()
        driver.find_element(By.XPATH, '//*[@id="DateOfFilingTo"]').send_keys(date_to)

        # Click Search
        driver.find_element(By.XPATH, '//*[@id="cmdSearch"]').click()
        time.sleep(2)

        # Check if results exist
        try:
            status_text = driver.find_element(By.XPATH, '//*[@id="tblStatus"]/tbody/tr/td').text
            ws.append([status_text]) 
            # Example: "Viewing Page 1 of 25 (494 RECORDS TOTAL)"
            
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

            wb.save("MD_Case_Search_Date_Range_Scrapping_Tool_Output.xlsx")

            current_page += 1
            if current_page <= total_pages:
                if not go_to_page(driver, current_page):
                    print(f"[!] Failed to navigate to page {current_page}")
                    break
        
        if records_scraped == total_records_reported:
            print(f"[{county}] ✅ Record count matches: {records_scraped} records scraped.")
        else:
            print(f"[{county}] ❌ Record count mismatch! Reported: {total_records_reported}, Scraped: {records_scraped}")

except:
    print(f"[!] Unexpected error")

# Final save and cleanup
wb.save("MD_Case_Search_Date_Range_Scrapping_Tool_Output.xlsx")
driver.quit()
print("✅ Done! Data saved to 'MD_Case_Search_Date_Range_Scrapping_Tool_Output.xlsx'")
root = tk.Tk()
root.withdraw()
root.attributes("-topmost",True)
messagebox.showinfo("Completed", "✅ Done! Data saved to 'MD_Case_Search_Date_Range_Scrapping_Tool_Output.xlsx'")
root.destroy()