import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Full path for Chrome user cache
cache_path = r"C:\Users\ERP3\Documents\Rick Apps\Selenium automation\selenium_cache"

# Set up Chrome options to use this cache
options = Options()
options.add_argument(f"user-data-dir={cache_path}")
# Optional: start browser maximized
options.add_argument("--start-maximized")

# Set up Chrome driver
driver = webdriver.Chrome(service=Service("./chromedriver.exe"), options=options)
wait = WebDriverWait(driver, 20)

# Read login credentials from pass.txt
try:
    with open("pass.txt", "r") as f:
        lines = f.read().splitlines()
        username = lines[0].strip()
        password = lines[1].strip()
except Exception as e:
    print(f"Failed to read pass.txt: {e}")
    driver.quit()
    exit()

# Open login page
driver.get("http://test.exactllyweb.com/Home")

try:
    # Wait for either username input or master button to detect login status
    if wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="validate-form"]/div[1]/input'))):
        print("[INFO] Not logged in, entering credentials...")
        
        driver.find_element(By.XPATH, '//*[@id="validate-form"]/div[1]/input').send_keys(username)
        driver.find_element(By.XPATH, '//*[@id="validate-form"]/div[2]/div/input').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="validate-form"]/div[5]/button[1]').click()
    else:
        print("[INFO] Already logged in.")

    # Wait for Master button
    master_btn_xpath = '/html/body/div/div[2]/section/div/div/div/div/div/div/div/div/a[1]/div'
    wait.until(EC.presence_of_element_located((By.XPATH, master_btn_xpath)))
    print("[INFO] Master button appeared, clicking...")
    driver.find_element(By.XPATH, master_btn_xpath).click()
    print("[SUCCESS] Login complete, Master button clicked.")

except Exception as e:
    print(f"[ERROR] Login process failed: {e}")

# Do not close browser for inspection
# driver.quit()
