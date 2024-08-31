import os
import tkinter as tk
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions as Options, ActionChains
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from sheetfu import SpreadsheetApp, Table
from google.oauth2 import service_account
import gspread

def get_google_sheet_data():
    # Load credentials and authenticate
    scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    credentials = service_account.Credentials.from_service_account_file(
        'mailer-400406-83227f4a1b2d.json', scopes=scope)
    gc = gspread.authorize(credentials)

    # Open the spreadsheet and select the sheet by name
    spreadsheet = gc.open_by_key('1i3v1YKO8R-doZ5ZyjBtJ-832WdZHJrt4tH6-VpTryeU')
    sheet = spreadsheet.worksheet('Sheet1')
    print(sheet.get_all_records())
    # Fetch all records from the sheet
    return sheet.get_all_records()


def update_google_sheet(transaction_no, error_message, status):
    # Load credentials and authenticate
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = service_account.Credentials.from_service_account_file(
        'mailer-400406-83227f4a1b2d.json', scopes=scope)
    gc = gspread.authorize(credentials)

    # Open the spreadsheet and select the sheet by name
    spreadsheet = gc.open_by_key('1i3v1YKO8R-doZ5ZyjBtJ-832WdZHJrt4tH6-VpTryeU')
    sheet = spreadsheet.worksheet('Sheet1')
    # Find the row where the transaction number matches
    cell = sheet.find(transaction_no)
    # Update the "Error" column for that row with the error message
    sheet.update_cell(cell.row, sheet.find("Error").col, error_message)
    sheet.update_cell(cell.row, sheet.find("Status").col, status)


def logistics(driver, tracking_no, po_no, row, status_label):
    wait = WebDriverWait(driver, 30)

    try:
        wait.until(ec.presence_of_element_located((By.XPATH, '//*[@id="Documentation"]/a')))

        # Click on the appropriate links to navigate
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH,
                                                                           '//*[@id="Documentation"]/div/span/div[2]/div/div/ul/li[2]/a'))
        time.sleep(2)
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH,
                                                                           '//*[@id="Documentation"]/div/span/div[2]/div/div/ul/li[2]/div/div/div[3]/ul/li/a'))
        time.sleep(5)

        # Enter tracking number and search
        trans = driver.find_element(By.XPATH,
                                    '//*[@id="content"]/form/div/div/div/div[2]/div/div[1]/div[2]/div/div/lss-dynamic-attr-text/div/input')
        trans.send_keys(tracking_no)
        driver.execute_script("arguments[0].scrollIntoView(true);", trans)
        time.sleep(1)
        search_button = driver.find_element(By.XPATH,
                                            '//*[@id="content"]/form/div/div/div/div[2]/div/div[4]/div/lss-dynamic-attr-button[1]/button')
        search_button.click()
        time.sleep(3)

        # Click on edit button
        edit_button = driver.find_element(By.XPATH, '//*[@id="actions"]/div/a[1]/span')
        edit_button.click()
        time.sleep(5)  # Add appropriate waiting here
        # Switch to the new tab
        driver.switch_to.window(driver.window_handles[-1])
        # Now navigate to the correct page using the tracking number
        # driver.get(f"https://lss.apllogistics.com/portal/#/SIP/sipdocupload?TRAN_TYPE=ETN&TRAN_NBR={tracking_no}")
        # time.sleep(5)  # Add appropriate waiting here
        # Find and click on the "Bulk Upload" tab
        try:
            bulk_upload_tab = wait.until(
                ec.element_to_be_clickable(
                    (By.XPATH, '//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/ul/li[2]/a')))
            bulk_upload_tab.click()
            print("clicked")
            # Dictionary mapping conditions to option values
            condition_to_option = {
                "INV": "CI - COMMERCIAL INVOICE",
                "IC": "IR - INSPECCION REPORT",
                "PL": "PL - PACKING LIST",
                "SUMMARY": "IS - INVOICE SUMMARY",
                "CHECKLIST": "DOCCK - DOCUMENT CHECKLIST",
                "CPSC": "CPSC - CONSUMER PRODUCT SAFETY COMMISSION CERT"
            }

            # Loop through conditions and click on options if conditions are met
            for condition, option_value in condition_to_option.items():
                if row[condition] == "✓":
                    try:
                        option = driver.find_element(By.XPATH,
                                                     f'//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div[1]/div/lss-duallistbox/div/table/tbody/tr[3]/td[1]/select/optgroup/option[text()="{option_value}"]')
                        option.click()
                        add = driver.find_element(By.XPATH,
                                                  '//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div[1]/div/lss-duallistbox/div/table/tbody/tr[3]/td[2]/div[3]/button')
                        add.click()
                        print(f"Clicked on option for condition {condition}: {option_value}")
                    except NoSuchElementException:
                        error_message = f"File types not found in APL."
                        update_google_sheet(tracking_no, error_message,
                                            "Error")  # Update the error column without proceeding to the browser
                        return  # Move to the next iteration of the loop
        except:
            print("Error: Failed to click on Bulk Upload tab")
        browse = driver.find_element(By.XPATH,
                                     '//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div[3]/div/lss-dynamic-attr-file/div/div/input')
        driver.execute_script("arguments[0].click();", browse)
        status_label.config(text="Clicked Browse button")
        status_label.update()  # Update the GUI to reflect the changes

        # # Upload files from 'final-copy-renamed' folder
        folder_path = fr"D:\GAP_MERGING_FILES(BOT)\\{po_no}-Combined"
        files_to_upload = os.listdir(folder_path)
        print(files_to_upload)
        for file_name in files_to_upload:
            # Check if the file corresponds to the transaction's Po No.
            if po_no in file_name:
                # Construct the file path directly using po_no
                file_path = os.path.join(folder_path, f"{po_no}-Merged.pdf")
                status_label.config(text=f"File found {file_name}")
                status_label.update()  # Update the GUI to reflect the changes
                # Provide the complete file path to the send_keys() method
                browse.send_keys(file_path)
                break
        time.sleep(1)
        status_label.config(text=f"Uploaded files")
        status_label.update()  # Update the GUI to reflect the changes
        save_btn = driver.find_element(By.XPATH,
                                       '//*[@id="content"]/form/div/div/div/div/div[3]/div[2]/div/div/div/div[2]/div[6]/div/lss-dynamic-attr-button[1]/button')
        time.sleep(1)
        ActionChains(driver).move_to_element(save_btn).click().perform()
        time.sleep(3)
        status = "Uploaded"
        update_google_sheet(tracking_no, "", status)  # Pass an empty error message for successful uploads
    except Exception as e:
        print(f"Error processing transaction {tracking_no}: {str(e)}")
        error_message = str(e)
        update_google_sheet(tracking_no, error_message, "Error")


def main():
    # Create the GUI window
    root = tk.Tk()
    root.geometry("400x100")
    root.title("API Doc Upload")  # Set the title of the window
    status_label = tk.Label(root, text="Processing...", padx=30, pady=15)
    status_label.pack()

    # login
    options = Options()
    # options.add_argument("--headless")
    # options.add_argument("--no-sandbox")
    # options.add_argument("--window-size=1280,720")
    # options.add_argument("--disable-gpu")
    options.add_experimental_option("detach", True)

    print("testing started")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)
    driver.get("https://lss.apllogistics.com/portal/#/home")
    username = wait.until(ec.presence_of_element_located((By.NAME, "username")))
    time.sleep(2)
    driver.execute_script("handleChangeLogin();")
    driver.execute_script("arguments[0].value = 'sunil.gopal@shahi.co.in';", username)
    driver.execute_script("arguments[0].value = 'Pwd_su12#';", driver.find_element(By.NAME, "password"))
    driver.execute_script("arguments[0].click();", driver.find_element(By.NAME, "submit"))

    # Get data from Google Sheet
    records = get_google_sheet_data()

    # Iterate through each record and process
    for record in records:
        tracking_no = record["Tracking No"]
        po_no = record["Po No"]
        logistics(driver, tracking_no, po_no, record, status_label)

    # Close the driver once all transactions are processed
    driver.quit()

    # Notify user that all transactions are processed
    status_label.config(text="All transactions processed.")
    status_label.update()

    # Keep the GUI window open
    root.mainloop()

