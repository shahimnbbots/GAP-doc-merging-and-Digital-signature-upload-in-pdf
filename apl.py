import os
import tkinter as tk
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions as Options, ActionChains
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from google.oauth2 import service_account
import gspread
import sys
import psutil


def get_google_sheet_data():
    scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    credentials = service_account.Credentials.from_service_account_file(
        'mailer-400406-83227f4a1b2d.json', scopes=scope)
    gc = gspread.authorize(credentials)
    spreadsheet = gc.open_by_key('1i3v1YKO8R-doZ5ZyjBtJ-832WdZHJrt4tH6-VpTryeU')
    sheet = spreadsheet.worksheet('Sheet1')
    return sheet.get_all_records()


def update_google_sheet(transaction_no, error_message, status):
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = service_account.Credentials.from_service_account_file(
        'mailer-400406-83227f4a1b2d.json', scopes=scope)
    gc = gspread.authorize(credentials)
    spreadsheet = gc.open_by_key('1i3v1YKO8R-doZ5ZyjBtJ-832WdZHJrt4tH6-VpTryeU')
    sheet = spreadsheet.worksheet('Sheet1')
    cell = sheet.find(transaction_no)
    sheet.update_cell(cell.row, sheet.find("Error").col, error_message)
    sheet.update_cell(cell.row, sheet.find("Status").col, status)


def get_chrome_pids():
    pids = []
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'] == 'chrome.exe':
            pids.append(proc.info['pid'])
    return pids


def kill_specific_chrome_processes(pids):
    for pid in pids:
        try:
            proc = psutil.Process(pid)
            proc.terminate()  # Or proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


def logistics(tracking_no, po_no, row, status_label):
    options = Options()
    # options.add_argument("--headless")
    # options.add_argument("--no-sandbox")
    # options.add_argument("--window-size=1280,720")
    # options.add_argument("--disable-gpu")

    initial_pids = get_chrome_pids()

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)
    try:
        driver.get("https://lss.apllogistics.com/portal/#/home")
        username = wait.until(ec.presence_of_element_located((By.NAME, "username")))
        time.sleep(2)
        driver.execute_script("handleChangeLogin();")
        driver.execute_script("arguments[0].value = 'sandesh.samson@shahi.co.in';", username)
        driver.execute_script("arguments[0].value = 'Booking@gap7';", driver.find_element(By.NAME, "password"))
        driver.execute_script("arguments[0].click();", driver.find_element(By.NAME, "submit"))
        time.sleep(5)
        wait.until(ec.presence_of_element_located((By.XPATH, '//*[@id="Documentation"]/a')))
        status_label.config(text=f"Logged in")
        status_label.update()

        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH,
                                                                           '//*[@id="Documentation"]/div/span/div[2]/div/div/ul/li[2]/a'))
        time.sleep(2)
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH,
                                                                           '//*[@id="Documentation"]/div/span/div[2]/div/div/ul/li[2]/div/div/div[3]/ul/li/a'))
        time.sleep(5)

        trans = driver.find_element(By.XPATH,
                                    '//*[@id="content"]/form/div/div/div/div[2]/div/div[1]/div[2]/div/div/lss-dynamic-attr-text/div/input')
        trans.send_keys(tracking_no)
        driver.execute_script("arguments[0].scrollIntoView(true);", trans)
        time.sleep(1)
        search_button = driver.find_element(By.XPATH,
                                            '//*[@id="content"]/form/div/div/div/div[2]/div/div[4]/div/lss-dynamic-attr-button[1]/button')
        search_button.click()
        time.sleep(3)

        edit_button = driver.find_element(By.XPATH, '//*[@id="actions"]/div/a[1]/span')
        edit_button.click()
        time.sleep(5)
        driver.switch_to.window(driver.window_handles[-1])

        try:
            bulk_upload_tab = wait.until(
                ec.element_to_be_clickable(
                    (By.XPATH, '//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/ul/li[2]/a')))
            bulk_upload_tab.click()

            condition_to_option = {
                "INV": "CI - COMMERCIAL INVOICE",
                "IC": "IR - INSPECCION REPORT",
                "PL": "PL - PACKING LIST",
                "SUMMARY": "IS - INVOICE SUMMARY",
                "CHECKLIST": "DOCCK - DOCUMENT CHECKLIST",
                "CPSC": "CPSC - CONSUMER PRODUCT SAFETY COMMISSION CERT"
            }

            for condition, option_value in condition_to_option.items():
                if row[condition] == "✓":
                    try:
                        option = driver.find_element(By.XPATH,
                                                     f'//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div[1]/div/lss-duallistbox/div/table/tbody/tr[3]/td[1]/select/optgroup/option[text()="{option_value}"]')
                        option.click()
                        add = driver.find_element(By.XPATH,
                                                  '//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div[1]/div/lss-duallistbox/div/table/tbody/tr[3]/td[2]/div[3]/button')
                        add.click()
                    except NoSuchElementException:
                        error_message = f"File types not found in APL."
                        update_google_sheet(tracking_no, error_message,
                                            "Error")
                        return
        except:
            print("Error: Failed to click on Bulk Upload tab")

        browse = driver.find_element(By.XPATH,
                                     '//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div[3]/div/lss-dynamic-attr-file/div/div/input')
        driver.execute_script("arguments[0].click();", browse)
        status_label.config(text="Clicked Browse button")
        status_label.update()

        folder_path = fr"D:\GAP_MERGING_FILES(BOT)\\{po_no}-Combined"
        files_to_upload = os.listdir(folder_path)
        for file_name in files_to_upload:
            if po_no in file_name:
                file_path = os.path.join(folder_path, f"{po_no}-Merged.pdf")
                status_label.config(text=f"File found {file_name}")
                status_label.update()
                browse.send_keys(file_path)
                break
        time.sleep(1)
        status_label.config(text=f"Uploaded files")
        status_label.update()
        save_btn = driver.find_element(By.XPATH,
                                       '//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div[6]/div/lss-dynamic-attr-button[1]/button')
        time.sleep(1)
        ActionChains(driver).move_to_element(save_btn).click().perform()
        time.sleep(3)
        status = "Uploaded"
        update_google_sheet(tracking_no, "", status)
    finally:
        driver.quit()
        new_pids = get_chrome_pids()
        # Kill only the new Chrome processes that were started by Selenium
        kill_specific_chrome_processes([pid for pid in new_pids if pid not in initial_pids])


def test():
    root = tk.Tk()
    root.geometry("400x100")
    root.title("API Doc Upload")
    status_label = tk.Label(root, text="Processing...", padx=30, pady=15)
    status_label.pack()
    sheet_data = get_google_sheet_data()

    for row in sheet_data:
        po_no = str(row["Po No."])
        tracking_no = str(row["Transaction No."])
        conditions_met = all(row[key] == "✓" for key in ["CHECKLIST", "INV", "PL", "IC"])
        status = row["Status"]
        if conditions_met and status != "Uploaded":
            try:
                status_label.config(text=f"Processing Transaction number: {tracking_no}")
                status_label.update()
                logistics(tracking_no, po_no, row, status_label)
            except:
                error_message = "Unable to login"
                update_google_sheet(tracking_no, error_message, "")
                sys.exit()
        elif not conditions_met:
            error_message = "Conditions are not met."
            update_google_sheet(tracking_no, error_message, "")
            print(f"Skipping logistics for transaction number {tracking_no} as conditions are not met.")
    root.destroy()


test()
