import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import logging
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import requests

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


CURRENT_VERSION = "1.0.0"
VERSION_URL = "https://raw.githubusercontent.com/gdscjec/python/main/version.txt"  

def check_for_update():
    try:
        response = requests.get(VERSION_URL)
        response.raise_for_status()
        latest_version = response.text.strip()
        if latest_version != CURRENT_VERSION:
            print(f"A new version of the bot is available (v{latest_version}). Please update to the latest version.")
            return False
        return True
    except Exception as e:
        logging.error(f"Error checking for update: {e}")
        return False

# Check for updates before proceeding
if not check_for_update():
    exit()

# The rest of your bot code
excel_file = 'mbg.xlsx'
data = pd.read_excel(excel_file)

if 'Google Review URL' not in data.columns:
    data['Google Review URL'] = ''
if 'Generated URL' not in data.columns:
    data['Generated URL'] = ''

driver_path = os.path.join(os.getcwd(), 'chromedriver')
driver = webdriver.Chrome()

review_link_generator_url = 'https://reviewlinkgenerator.com/'
qr_code_generator_url = 'https://create.reviewus.in/Self_Generation'

for index, row in data.iterrows():
    logging.info(f"Processing row {index + 1} for review link generation")
    driver.get(review_link_generator_url)

    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.NAME, 'places_autocomplete'))
        )

        business_name_input = driver.find_element(By.NAME, 'places_autocomplete')
        business_name_input.send_keys(row['Business Info'])
        
        time.sleep(2)

        dropdown_first_item = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.pac-item'))
        )
        dropdown_first_item.click()

        test_link_button = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'test_place_link'))
        )

        time.sleep(4)

        generated_url = test_link_button.get_attribute('href')
        
        if generated_url and generated_url != 'https://reviewlinkgenerator.com/#':
            data.at[index, 'Google Review URL'] = generated_url
            logging.info(f"Captured URL: {generated_url}")
        else:
            logging.error(f"No valid URL found in href attribute for row {index + 1}")

        time.sleep(2)
    except TimeoutException as e:
        logging.error(f"Timeout while processing row {index + 1}: {e}")
        with open(f'error_row_{index + 1}.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        screenshot_path = os.path.join(os.getcwd(), f'error_row_{index + 1}.png')
        driver.save_screenshot(screenshot_path)
        logging.info(f"Saved screenshot to {screenshot_path}")
    except NoSuchElementException as e:
        logging.error(f"Element not found while processing row {index + 1}: {e}")
    except Exception as e:
        logging.error(f"Error processing row {index + 1}: {e}")

for index, row in data.iterrows():
    logging.info(f"Processing row {index + 1} for QR code generation")
    driver.get(qr_code_generator_url)
    
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.NAME, 'business_name'))
        )
        driver.find_element(By.NAME, 'business_name').send_keys(row['Business Name'])
        driver.find_element(By.NAME, 'owner_email').send_keys(row['Owner Email'])
        driver.find_element(By.NAME, 'review_url').send_keys(row['Google Review URL'])
        driver.find_element(By.NAME, 'phone_no').send_keys(row['Enter Contact No'])
        driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()

        WebDriverWait(driver, 20).until(
            EC.url_changes(qr_code_generator_url)
        )
        generated_url = driver.current_url
        data.at[index, 'Generated URL'] = generated_url
        logging.info(f"Captured URL: {generated_url}")

        time.sleep(5)
    except Exception as e:
        logging.error(f"Error processing row {index + 1}: {e}")
driver.quit()
data.to_excel(excel_file, index=False)
logging.info("Process completed and Excel file saved.")
print("Developed By: Ankit Pawar")
