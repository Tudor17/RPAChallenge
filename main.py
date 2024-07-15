from datetime import datetime
from dateutil.relativedelta import relativedelta
import time
import re
import os
import openpyxl
from RPA.Browser.Selenium import Selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Function to extract monetary values from text
def extract_money(text):
    money_pattern = r'\$[\d,.]+|\d+ dollars|\d+ USD'
    matches = re.findall(money_pattern, text)
    return len(matches) > 0

# Initialize Browser
browser = Selenium()

try:
    # Open the website
    browser.open_available_browser("https://www.latimes.com/", maximized=True)

    # Enter search phrase (e.g., "Donald Trump") in the search field
    search_phrase = 'Donald Trump'
    section = "business"
    selection = 2

    if selection <= 0:
        selection = 1

    target_date = datetime.now() - relativedelta(months=selection - 1)
    target_date = target_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

    WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/ps-header/header/div[2]/button")))
    browser.click_element("xpath:/html/body/ps-header/header/div[2]/button")
    
    search_input = browser.driver.find_element(By.XPATH, '/html/body/ps-header/header/div[2]/div[2]/form/label/input')
    search_input.send_keys(search_phrase)
    search_input.submit()

    # Wait for search results
    WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "search-results-module")))

    # Choose the latest (newest) news
    browser.select_from_list_by_value("xpath:/html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[1]/div[2]/div/label/select", "1")

    # Filter by section
    WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "checkbox-input-label")))
    #results = browser.driver.find_elements(By.CLASS_NAME, 'checkbox-input-label')

    #for result in results:
    #    WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, browser.get_element_attribute(result, ""))))

    #matching_element = next((element for element in results if section.lower() in browser.get_text(element).lower()), None)

    #if matching_element:
    #    browser.click_element(matching_element)

    time.sleep(2)

    last_page = False
    counter = 1

    while last_page == False:
        # Get values: title, date, description
        results = browser.driver.find_elements(By.CLASS_NAME, 'promo-wrapper')
        for result in results:

            WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "promo-timestamp")))
            date = datetime.fromtimestamp(float(browser.get_element_attribute(result.find_element(By.CLASS_NAME, 'promo-timestamp'), 'data-timestamp'))/1000)

            if date < target_date:
                last_page = True
                break

            WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "promo-title")))
            title = browser.get_text(result.find_element(By.CLASS_NAME, 'promo-title'))
            WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "promo-description")))
            description = browser.get_text(result.find_element(By.CLASS_NAME, 'promo-description'))
            WebDriverWait(browser.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "promo-media")))
            image_filename = browser.capture_element_screenshot(result.find_element(By.CLASS_NAME, 'promo-media'), str(counter)+'.png')
            title_count = title.lower().count(search_phrase.lower())
            description_count = description.lower().count(search_phrase.lower())
            contains_money = extract_money(title) or extract_money(description)

            # Store data in Excel file
            excel_file = "news_data.xlsx"
            if not os.path.exists(excel_file):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(["Title", "Date", "Description", "Picture Filename", "Search Phrase Count (Title)", "Search Phrase Count (Description)", "Contains Money"])
            else:
                wb = openpyxl.load_workbook(excel_file)
                ws = wb.active
        
            ws.append([title, date, description, image_filename, title_count, description_count, contains_money])
            wb.save(excel_file)
            counter = counter + 1
        try:
            WebDriverWait(browser.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[2]/div[3]/a"))).click()
        except:
            last_page = True

finally:
    # Close browser and cleanup
    browser.close_all_browsers()