import tkinter as tk
from tkinter import scrolledtext
from threading import Thread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
import pyautogui
import csv
import os
from selenium.webdriver.chrome.options import Options

# Define automation class
class WebAutomation:
    def __init__(self, log_callback):
        self.driver = None
        self.running = False
        self.stop_requested = False
        self.log = log_callback
        self.login_url = "http://test.website.com/Home"
        self.target_product_codes = []
        self.processed_codes = []
        self.resume_from_code = None

        try:
            with open("pass.txt", "r") as f:
                lines = f.read().splitlines()
                self.username = lines[0].strip()
                self.password = lines[1].strip()
        except Exception as e:
            self.username = ""
            self.password = ""
            self.log(f"Failed to read credentials: {e}")

        try:
            with open("items.csv", newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader)
                self.target_product_codes = [row[1].strip() for row in reader if row[1].strip()]
        except Exception as e:
            self.log(f"Failed to read items.csv: {e}")

        try:
            with open("resume.txt", "r") as f:
                self.resume_from_code = f.read().strip()
        except:
            self.resume_from_code = None

    def save_processed_codes(self):
        with open("processed_rows.csv", "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["ProductCode", "Timestamp"])
            for entry in self.processed_codes:
                writer.writerow(entry)

    def highlight(self, element):
        self.driver.execute_script("arguments[0].style.border='3px solid red'", element)

    def start(self):
        self.stop_requested = False

        # Chrome user cache setup
        cache_path = r"C:\\Users\\ERP3\\Documents\\Rick Apps\\Selenium automation\\selenium_cache"
        options = Options()
        options.add_argument(f"user-data-dir={cache_path}")
        options.add_argument("--start-maximized")

        self.driver = webdriver.Chrome(service=Service('./chromedriver.exe'), options=options)
        self.driver.get(self.login_url)

        try:
            self.log("Filling username...")
            try:
                self.driver.find_element(By.XPATH, '//*[@id="validate-form"]/div[1]/input').send_keys(self.username)
                self.driver.find_element(By.XPATH, '//*[@id="validate-form"]/div[2]/div/input').send_keys(self.password)
                self.log("Clicking login button...")
                self.driver.find_element(By.XPATH, '//*[@id="validate-form"]/div[5]/button[1]').click()
                time.sleep(3)
            except:
                self.log("Already logged in or session reused.")

            self.log("Clicking master button...")
            self.driver.find_element(By.XPATH, '/html/body/div/div[2]/section/div/div/div/div/div/div/div/div/a[1]/div').click()
            time.sleep(3)

            self.log("Hovering on Optical Master...")
            menu = self.driver.find_element(By.XPATH, '//*[@id="cd-primary-nav"]/li[5]/a')
            ActionChains(self.driver).move_to_element(menu).perform()
            time.sleep(1)

            self.log("Hovering on Item Master...")
            submenu = self.driver.find_element(By.XPATH, '//*[@id="cd-primary-nav"]/li[5]/ul/li[2]/a')
            ActionChains(self.driver).move_to_element(submenu).perform()
            time.sleep(1)

            frame_link_xpath = '//*[@id="cd-primary-nav"]/li[5]/ul/li[2]/ul/li[1]/a'
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, frame_link_xpath)))
            try:
                self.driver.find_element(By.XPATH, frame_link_xpath).click()
                self.log("Clicked Frame link.")
            except Exception:
                frame_element = self.driver.find_element(By.XPATH, frame_link_xpath)
                self.driver.execute_script("arguments[0].click();", frame_element)

            time.sleep(3)
            self.driver.switch_to.window(self.driver.window_handles[-1])
            time.sleep(10)  # allow manual setting to view all entries

            resume_found = self.resume_from_code is None

            while not self.stop_requested:
                try:
                    code_elements = self.driver.find_elements(By.XPATH, "//*[@id][contains(@id, '-uiGrid-0006-cell')]/div/u")
                    codes_on_page = [el.text.strip() for el in code_elements]

                    for i in range(len(codes_on_page)):
                        code = codes_on_page[i]

                        if not resume_found:
                            if code == self.resume_from_code:
                                resume_found = True
                            else:
                                continue

                        if code in self.target_product_codes and code not in [row[0] for row in self.processed_codes]:
                            self.log(f"Processing Product Code: {code} (row {i})")
                            edit_xpath = f"//*[@id][contains(@id,'-uiGrid-0005-cell') and contains(@id,'-{i}-')]/button[3]"
                            edit_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, edit_xpath)))
                            self.highlight(edit_button)
                            edit_button.click()
                            time.sleep(5)

                            self.log("Checking checkbox...")
                            try:
                                WebDriverWait(self.driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH, '//*[@id="frm_OPTICAL_ITEM"]/div/fieldset[3]/div/div/div/div/table/tbody/tr/td[2]/input'))
                                )
                                checkboxes = self.driver.find_elements(By.XPATH, '//*[@id="frm_OPTICAL_ITEM"]/div/fieldset[3]/div/div/div/div/table/tbody/tr/td[2]/input')
                            except:
                                checkboxes = []

                            if not checkboxes:
                                self.log("No checkboxes found.")
                            else:
                                for cb in checkboxes:
                                    try:
                                        self.driver.execute_script("arguments[0].scrollIntoView(true);", cb)
                                        if not cb.is_selected():
                                            self.highlight(cb)
                                            cb.click()
                                            time.sleep(0.5)
                                    except Exception as e:
                                        self.log(f"Checkbox interaction failed: {e}")

                            self.log("Clicking save button...")
                            self.driver.find_element(By.XPATH, '/html/body/div/div[2]/section/div[2]/div[1]/div/div[2]/ul/li[1]/a/span/i').click()
                            time.sleep(1)

                            self.log("Clicking final save in dialog...")
                            self.driver.find_element(By.XPATH, '/html/body/div[3]/div/div/form/div/div[3]/button[2]').click()
                            time.sleep(5)

                            try:
                                self.log("Clicking optional OK...")
                                self.driver.find_element(By.XPATH, '//*[@id="frm_OPTICAL_ITEM"]/div/div[2]/div/div[3]/button').click()
                                time.sleep(1)
                            except:
                                self.log("OK button not present.")

                            self.driver.execute_script("window.onbeforeunload = null;")
                            self.log("Navigating back to list view...")
                            self.driver.back()
                            time.sleep(5)

                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            self.processed_codes.append([code, timestamp])
                            self.save_processed_codes()

                    self.log("Clicking Next Page...")
                    next_btn = self.driver.find_element(By.XPATH, '/html/body/div/div[2]/section/div[2]/div[2]/div/div/div/div[3]/div/div[1]/button[2]')
                    next_btn.click()
                    time.sleep(5)

                except Exception as e:
                    self.log(f"Item failed or not found: {e}")
                    break

        except Exception as e:
            self.log(f"Fatal Error: {e}")

    def stop(self):
        self.stop_requested = True
        self.running = False
        self.log("Stop signal received. Browser will remain open.")

# GUI setup
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Web Automation")

        self.start_button = tk.Button(root, text="Start", command=self.start_thread)
        self.start_button.pack(pady=5)

        self.stop_button = tk.Button(root, text="Stop", command=self.stop)
        self.stop_button.pack(pady=5)

        self.log_box = scrolledtext.ScrolledText(root, width=80, height=20)
        self.log_box.pack(padx=10, pady=10)

        self.automation = WebAutomation(self.log)

    def log(self, message):
        timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        self.log_box.insert(tk.END, f"{timestamp} {message}\n")
        self.log_box.see(tk.END)

    def start_thread(self):
        self.log("Starting automation using Product Code match...")
        Thread(target=self.automation.start).start()

    def stop(self):
        self.automation.stop()

# Run the app
if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
