import os.path
import shutil
import time
from tkinter import filedialog, messagebox
import gspread
import pandas as pd
import os.path
import PyPDF2
import tkinter as tk
from google.oauth2 import service_account
from gspread.exceptions import APIError
from sheetfu import SpreadsheetApp, Table


def add_data(data, po_number):
    print("Adding data to Google Sheet:", data)
    # Load credentials and authenticate
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = service_account.Credentials.from_service_account_file(
        'mailer-400406-83227f4a1b2d.json', scopes=scope)
    gc = gspread.authorize(credentials)

    # Open the spreadsheet and select the sheet by name
    spreadsheet = gc.open_by_key('1i3v1YKO8R-doZ5ZyjBtJ-832WdZHJrt4tH6-VpTryeU')
    sheet = spreadsheet.worksheet('Sheet1')

    # Get all values from the first row (column headers)
    headers = sheet.row_values(1)

    # Find the cell where the transaction number matches
    cell = None
    try:
        cell = sheet.find(po_number)
    except gspread.exceptions.CellNotFound:
        pass

    if cell:
        # If the cell is found, update the existing row
        row_index = cell.row
    else:
        # If the cell is not found, create a new row
        row_index = len(sheet.get_all_values()) + 1
        sheet.insert_row([None] * len(headers), index=row_index)

    # Attempt to add data to the Google Sheet
    retries = 0
    max_retries = 3
    while retries < max_retries:
        try:
            # Attempt to add data
            # Update the data for the relevant columns
            for col_name in data:
                # Check if the column name exists in the headers
                if col_name in headers:
                    col_index = headers.index(col_name) + 1  # Adjust index to 1-based
                    sheet.update_cell(row_index, col_index, data[col_name])
                else:
                    print(f"Column '{col_name}' not found in Google Sheet.")

            print("Data updated in Google Sheet.")
            break  # Successful, exit the loop

        except APIError as e:
            # Handle quota exceeded error here
            if e.response.status_code == 429:  # Quota exceeded error
                print("Quota exceeded. Retrying after 60 seconds...")
                time.sleep(60)  # Wait for 60 seconds
                retries += 1
            else:
                print("An unexpected error occurred:", e)
                break  # Exit the loop for unexpected errors

    else:
        print("Max retries reached. Failed to add data to Google Sheet.")


def merging(po_number, tracking_no):
    status_label.config(text=f"Processing....")
    status_label.update()  # Update the GUI to reflect the changes
    # Define paths
    folder = 'D:\\GAP_MERGING_FILES(BOT)'
    # Create folder for the specific PO number if it doesn't exist
    po_folder_path = os.path.join(folder, f"{po_number}-Combined")
    if not os.path.exists(po_folder_path):
        os.makedirs(po_folder_path)
    # Move the files to the PO folder
    for filename in os.listdir(folder):
        if po_number in filename:
            file_path = os.path.join(folder, filename)
            # If it's an "IC" file, copy it instead of moving
            if filename.endswith("IC.pdf"):
                shutil.copy(file_path, po_folder_path)
            else:
                shutil.move(file_path, po_folder_path)

    # Create a PDF writer object
    pdf_writer = PyPDF2.PdfWriter()

    # Loop through each PDF file in the folder
    for filename in os.listdir(po_folder_path):
        if filename.endswith(".pdf"):
            # Open each PDF file in read-binary mode
            with open(os.path.join(po_folder_path, filename), 'rb') as file:
                # Create a PDF reader object
                pdf_reader = PyPDF2.PdfReader(file)
                # Loop through each page in the PDF and add it to the writer
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    pdf_writer.add_page(page)

    # Write the merged PDF to the output file
    merged_pdf_path = os.path.join(po_folder_path, f"{po_number}-Merged.pdf")
    with open(merged_pdf_path, 'wb') as output_file:
        pdf_writer.write(output_file)

    print("Merged PDF saved at:", merged_pdf_path)

    # Create a folder for removed files
    removed_files_folder = os.path.join(po_folder_path, f"{po_number}")
    if not os.path.exists(removed_files_folder):
        os.makedirs(removed_files_folder)
    # Move the files to the removed files folder before removing them
    for filename in os.listdir(po_folder_path):
        if filename.endswith(".pdf") and not filename.startswith(f"{po_number}-Merged"):
            shutil.move(os.path.join(po_folder_path, filename), os.path.join(removed_files_folder, filename))
            print("Moved to Removed_Files:", filename)
    # Initialize data dictionary
    data = {
        "Po No.": po_number,
        "Transaction No.": tracking_no,
    }
    print("Files in the folder:")
    #
    # Suffixes to check
    suffixes = ["INV", "PL", "CPSC", "SUMMARY", "IC", "CHECKLIST"]
    mandatory_suffixes = ["INV", "PL", "IC", "CHECKLIST"]
    # Search for files with suffixes and update data accordingly
    all_mandatory_files_found = all(
        any(filename.endswith(suffix + ".pdf") for filename in os.listdir(removed_files_folder))
        or any(filename.endswith("CH.pdf") for filename in os.listdir(removed_files_folder))
        if suffix == "CHECKLIST" else
        any(filename.endswith(suffix + ".pdf") for filename in os.listdir(removed_files_folder))
        for suffix in mandatory_suffixes)

    if all_mandatory_files_found:
        for suffix in suffixes:
            if suffix == "CHECKLIST":
                # For CHECKLIST, check both CH and CHECKLIST
                file_found = any(filename.endswith("CH.pdf") or filename.endswith("CHECKLIST.pdf") for filename in
                                 os.listdir(removed_files_folder))
            else:
                file_found = any(filename.endswith(suffix + ".pdf") for filename in os.listdir(removed_files_folder))

            data[suffix] = "✓" if file_found else "❌"

        print("Updated data:", data)
        data["Error"] = f""
        print("Error:", data["Error"])
        add_data(data, po_number)
    else:
        # Add an error message if mandatory files are not found
        data["Error"] = f"Mandatory files are not found for Po: {po_number}"
        print("Error:", data["Error"])
        add_data(data, po_number)


def test():
    # Choose the file using file dialog
    filetypes = (('excel files', '*.xls'), ('excel files', '*.xlsx'), ('excel files', '*.ods'))
    filename = filedialog.askopenfilename(
        title='Open excel',
        initialdir="/",
        filetypes=filetypes
    )

    df = pd.read_excel(filename)

    # Print column names to verify they are correct
    print(df.columns)

    po_number_a = df['PO#'].tolist()
    tracking_no_a = df['TRNX/BKG Number'].tolist()
    for i in range(len(po_number_a)):
        po_number = po_number_a[i]
        tracking_no = tracking_no_a[i]
        status_label.config(text=f"Processing {po_number}")
        status_label.update()  # Update the GUI to reflect the changes
        merging(po_number, tracking_no)
        time.sleep(10)
    status_label.config(text=f"Completed")
    status_label.update()  # Update the GUI to reflect the changes
    time.sleep(2)
    root.destroy()


root = tk.Tk()
root.geometry("500x100")
root.title("GAP_Doc Merging")  # Set the title of the window
status_label = tk.Label(root, text="Started...", padx=30, pady=15)
status_label.pack()
root.after(100, test)  # Call the test function after 100 milliseconds
root.mainloop()
