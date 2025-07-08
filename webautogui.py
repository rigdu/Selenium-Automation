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
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import json
import csv
from datetime import datetime

class WebAutomation:
    def __init__(self, log_callback):
        self.driver = None
        self.running = False
        self.stop_requested = False
        self.log = log_callback
        self.target_product_codes = []
        self.processed_codes = []
        self.resume_from_code = None

        try:
            with open("config.json", "r") as f:
                config = json.load(f)
                self.cache_path = config.get("cache_path", "")
                self.login_url = config.get("url", "")
                self.username = config.get("username", "")
                self.password = config.get("password", "")
                self.resume_from_code = config.get("resume_from_code", None)
            self.log("Configuration loaded successfully from config.json")
        except Exception as e:
            self.log(f"Failed to read config.json: {e}")
            self.cache_path = ""
            self.login_url = ""
            self.username = ""
            self.password = ""
            self.resume_from_code = None

        try:
            with open("items.csv", newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader)  # Skip header
                self.target_product_codes = [row[1].strip() for row in reader if row[1].strip()]
            self.log(f"Loaded {len(self.target_product_codes)} product codes from items.csv")
        except Exception as e:
            self.log(f"Failed to read items.csv: {e}")

        if self.resume_from_code and self.resume_from_code not in self.target_product_codes:
            self.log(f"Invalid resume_from_code '{self.resume_from_code}' in config.json; starting from beginning")
            self.resume_from_code = None
        elif self.resume_from_code:
            self.log(f"Resuming from product code: {self.resume_from_code}")
        else:
            self.log("No resume_from_code in config.json; starting from beginning")

    def save_processed_codes(self):
        try:
            with open("processed_rows.csv", "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["ProductCode", "Timestamp"])
                for entry in self.processed_codes:
                    writer.writerow(entry)
            self.log("Saved processed codes to processed_rows.csv")
        except Exception as e:
            self.log(f"Failed to save processed codes: {e}")

    def highlight(self, element):
        try:
            self.driver.execute_script("arguments[0].style.border='3px solid red'", element)
            self.log("Highlighted element")
        except Exception as e:
            self.log(f"Failed to highlight element: {e}")

    def start(self):
        self.stop_requested = False
        self.running = True

        options = webdriver.ChromeOptions()
        options.add_argument(f"user-data-dir={self.cache_path}")
        options.add_argument("--start-maximized")

        try:
            self.driver = webdriver.Chrome(service=Service('./chromedriver.exe'), options=options)
            self.log("Browser initialized")
            self.driver.get(self.login_url)
            self.log(f"Navigated to {self.login_url}")

            try:
                username_field = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="validate-form"]/div[1]/input'))
                )
                username_field.send_keys(self.username)
                self.log("Filled username")
                password_field = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="validate-form"]/div[2]/div/input'))
                )
                password_field.send_keys(self.password)
                self.log("Filled password")
                login_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, '#validate-form > div.forgot_area > button:nth-child(1)'))
                )
                self.highlight(login_button)
                login_button.click()
                self.log("Clicked login button")
            except TimeoutException:
                self.log("Already logged in or session reused")
            except NoSuchElementException:
                self.log("Login button not found with provided selector")
                return

            try:
                master_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[2]/section/div/div/div/div/div/div/div/div/a[1]/div'))
                )
                self.highlight(master_button)
                master_button.click()
                self.log("Clicked master button")
            except TimeoutException:
                self.log("Master button not found")
                return

            try:
                menu = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="cd-primary-nav"]/li[5]/a'))
                )
                ActionChains(self.driver).move_to_element(menu).perform()
                self.log("Hovered on Optical Master")
            except TimeoutException:
                self.log("Optical Master menu not found")
                return

            try:
                submenu = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="cd-primary-nav"]/li[5]/ul/li[2]/a'))
                )
                ActionChains(self.driver).move_to_element(submenu).perform()
                self.log("Hovered on Item Master")
            except TimeoutException:
                self.log("Item Master submenu not found")
                return

            try:
                frame_link = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="cd-primary-nav"]/li[5]/ul/li[2]/ul/li[1]/a'))
                )
                self.highlight(frame_link)
                self.driver.execute_script("arguments[0].click();", frame_link)
                self.log("Clicked Frame link")
            except TimeoutException:
                self.log("Frame link not found")
                return

            try:
                WebDriverWait(self.driver, 10).until(EC.number_of_windows_to_be(2))
                self.driver.switch_to.window(self.driver.window_handles[-1])
                self.log("Switched to new window")
            except TimeoutException:
                self.log("New window not opened")
                return

            try:
                dropdown = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//select[@ng-model="grid.options.paginationPageSize"]'))
                )
                dropdown.click()
                all_option = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//select[@ng-model="grid.options.paginationPageSize"]//option[contains(text(),"All")]'))
                )
                self.highlight(all_option)
                all_option.click()
                self.log("Set dropdown to All")
                dropdown.send_keys(Keys.ESCAPE)
            except TimeoutException:
                self.log("Could not set dropdown to All")

            try:
                code_elements = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[id*='-uiGrid-0006-cell'] div u"))
                )
                codes_on_page = [el.text.strip() for el in code_elements]
                self.log(f"Found {len(codes_on_page)} product codes on page")

                resume_found = self.resume_from_code is None

                for i, code in enumerate(codes_on_page):
                    if self.stop_requested or not self.running:
                        self.log("Stop requested, exiting loop")
                        break

                    if not resume_found:
                        if code == self.resume_from_code:
                            resume_found = True
                            self.log(f"Resumed at product code: {code}")
                        continue

                    if code in self.target_product_codes and code not in [row[0] for row in self.processed_codes]:
                        self.log(f"Processing Product Code: {code} (row {i})")
                        try:
                            edit_button = WebDriverWait(self.driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, f"[id*='-{i}-uiGrid-0005-cell'] > button:nth-child(3)"))
                            )
                            self.highlight(edit_button)
                            edit_button.click()
                            self.log("Clicked edit button")
                        except TimeoutException:
                            self.log(f"Edit button not clickable for code {code}")
                            continue
                        except NoSuchElementException:
                            self.log(f"Edit button not found for code {code}")
                            continue

                        try:
                            WebDriverWait(self.driver, 20).until(
                                EC.presence_of_element_located((By.XPATH, '//*[@id="frm_OPTICAL_ITEM"]/div/fieldset[3]/div/div/div/div/table/tbody/tr/td[2]/input'))
                            )
                            checkboxes = self.driver.find_elements(By.XPATH, '//*[@id="frm_OPTICAL_ITEM"]/div/fieldset[3]/div/div/div/div/table/tbody/tr/td[2]/input')
                            if not checkboxes:
                                self.log("No checkboxes found")
                            else:
                                for cb in checkboxes:
                                    try:
                                        self.driver.execute_script("arguments[0].scrollIntoView(true);", cb)
                                        if not cb.is_selected():
                                            self.highlight(cb)
                                            cb.click()
                                            self.log("Checked checkbox")
                                    except Exception as e:
                                        self.log(f"Checkbox interaction failed: {e}")
                        except TimeoutException:
                            self.log("Checkbox section not loaded in time")

                        try:
                            save_button = WebDriverWait(self.driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[2]/section/div[2]/div[1]/div/div[2]/ul/li[1]/a/span/i'))
                            )
                            self.highlight(save_button)
                            save_button.click()
                            self.log("Clicked save button")
                            final_save = WebDriverWait(self.driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div/form/div/div[3]/button[2]'))
                            )
                            self.highlight(final_save)
                            final_save.click()
                            self.log("Clicked final save in dialog")
                        except TimeoutException:
                            self.log("Save button or dialog not found")
                            continue

                        try:
                            ok_button = WebDriverWait(self.driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, '//*[@id="frm_OPTICAL_ITEM"]/div/div[2]/div/div[3]/button'))
                            )
                            self.highlight(ok_button)
                            ok_button.click()
                            self.log("Clicked optional OK button")
                        except TimeoutException:
                            self.log("OK button not present")

                        self.driver.execute_script("window.onbeforeunload = null;")
                        self.driver.back()
                        self.log("Navigated back to list view")

                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        self.processed_codes.append([code, timestamp])
                        self.save_processed_codes()
            except Exception as e:
                self.log(f"Page processing failed: {e}")
        except Exception as e:
            # Catch all exceptions during setup (e.g., driver initialization, navigation)
            if isinstance(e, WebDriverException):
                self.log(f"Automation error (WebDriver): {str(e)}")
            else:
                self.log(f"Setup error: {str(e)}")

    def stop(self):
        self.stop_requested = True
        self.running = False
        self.log("Automation stopped, browser remains open")

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

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
