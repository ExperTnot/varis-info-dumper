import os
import re
import time
import sys
import re
from docx import Document
import openpyxl
from tkinter import Tk, Button, Label, Frame
import threading

CONFIG_FILE_DOCX = "config.txt"
CONFIG_FILE_XLSX = "config_xlsx.txt"

#Threading
def gui_thread(lines):
    r = Tk()
    r.title("Copy Lines to Clipboard")

    def copy_to_clipboard(line_to_copy):
        r.clipboard_clear()
        r.clipboard_append(line_to_copy)
        r.update()

    def close_window():
        r.quit()
        r.destroy()
    
    background_colors = ["white", "lightgrey", "darkgrey"]  # Define background colors
    row_num = 0  # Initialize row number

    for line_num, line in enumerate(lines, start=1):
        line = line.strip()
        frame = Frame(r, bg=background_colors[row_num % len(background_colors)])
        label_text = f"{line_num}. {line}"
        label = Label(frame, text=label_text, bg=background_colors[row_num % len(background_colors)])
        copy_button = Button(frame, text=f"Copy {line_num}", command=lambda line=line: copy_to_clipboard(line))
        label.pack(side="left")
        copy_button.pack(side="right")
        frame.pack(fill="both", expand=True)
        row_num += 1

    close_button = Button(r, text="Close", command=close_window)
    close_button.pack()

    r.protocol("WM_DELETE_WINDOW", close_window)
    
    window_height = len(lines) * 50  # Adjust the multiplier as needed
    r.geometry(f"200x{window_height}") # Adjust the width as needed
    r.mainloop()

# Function to extract data from a Word document
def extract_word_data(docx_file):
    document = Document(docx_file)
    data = []

    for paragraph in document.paragraphs:
        text = paragraph.text

        # Use regular expression to find HP numbers
        hp_numbers = re.findall(r'HP:\d{7}', text)

        if hp_numbers:
            data.extend(hp_numbers)

    return data

# Function to open a file with a 1-second delay
def open_file_with_delay(file_path):
    time.sleep(1)
    os.system(f"start {file_path}")

# Function to get the folder path from the user and save it to the configuration file
def get_folder_path():
    folder_path = input("Enter the folder path where your .docx files are located: ")
    with open(CONFIG_FILE_DOCX, "w") as config_file:
        config_file.write(folder_path)
    return folder_path

# Function to read the folder path from the configuration file
def read_folder_path():
    if os.path.exists(CONFIG_FILE_DOCX):
        with open(CONFIG_FILE_DOCX, "r") as config_file:
            return config_file.read()
    else:
        return None

# Custom exception for stopping the program
class StopProgramException(Exception):
    def __init__(self, message, value=None):
        super().__init__(message)
        self.value = value

# Function to search for a 4-digit number in an Excel sheet and extract data from a specific column (Column C -> Column I)
def search_excel_and_extract_data(excel_file, search_value):
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    extracted_data = None

    # Iterate through rows in Column C
    for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        cell_value = str(row[2])  # Column C is the third column (index 2)

        if search_value in cell_value:
            # Extract data from Column I (if available)
            column_i_value = None
            if len(row) > 8:  # Check if there are enough columns in the current row
                column_i_value = row[8]  # Column I (index 8)

            extracted_data = (cell_value, column_i_value)

            # Debug print after finding a matching cell in Column C
            print(f"Found '{search_value}' in cell '{cell_value}', extracted Column I value: '{column_i_value}'")

            # Search for "Vater" or "Mutter" in the next row
            next_row_number = row_number + 1  # Row number of the next row

            if next_row_number <= sheet.max_row:
                next_row = sheet.cell(row=next_row_number, column=9).value  # Column I in the next row

                if next_row in ["Vater", "Mutter"]:
                    next_row_c_value = None
                    
                    next_row_c_value = sheet.cell(row=next_row_number, column=4).value  # Column C in the next row

                    # Add "Vater" or "Mutter" and their corresponding Column C to extracted_data
                    extracted_data += (next_row, next_row_c_value)
                    
                    # Debug print after finding "Vater" or "Mutter" in the next row
                    print(f"Found '{next_row}' in the next row, Column C: '{next_row_c_value}'")

                    # Search for "Vater" or "Mutter" in the row after the next row
                    next_next_row_number = next_row_number + 1  # Row number of the row after the next row

                    if next_next_row_number <= sheet.max_row:
                        next_next_row = sheet.cell(row=next_next_row_number, column=9).value  # Column I in the row after the next row

                        if next_next_row in ["Vater", "Mutter"]:
                            next_next_row_c_value = None
                            
                            next_next_row_c_value = sheet.cell(row=next_next_row_number, column=4).value  # Column C in the row after the next row

                            # Add "Vater" or "Mutter" and their corresponding Column C to extracted_data
                            extracted_data += (next_next_row, next_next_row_c_value)
                            
                            # Debug print after finding "Vater" or "Mutter" in the row after the next row
                            print(f"Found '{next_next_row}' in the row after the next row, Column C: '{next_next_row_c_value}'")

                            # Break out of the loop
                            break
                        # Break out of the loop if "Vater" or "Mutter" is not found in the second next row
                        else:
                            break
                # Break out of the loop if "Vater" or "Mutter" is not found in the next row
                else:
                    break

    wb.close()
    return extracted_data

# Function to get the Excel file path from the user and save it to the configuration file
def get_excel_file_path():
    excel_file_path = input("Enter the path to the Excel sheet (xlsx file): ")
    with open(CONFIG_FILE_XLSX, "w") as config_file:
        config_file.write(excel_file_path)
    return excel_file_path

# Function to read the Excel file path from the configuration file
def read_excel_file_path():
    if os.path.exists(CONFIG_FILE_XLSX):
        with open(CONFIG_FILE_XLSX, "r") as config_file:
            return config_file.read()
    else:
        return None

# Function to add data to a text file
def add_data_to_text_file(file_path, extracted_data):
    with open(file_path, "a") as file:
        if extracted_data is not None:
            for item in extracted_data[1:]:  # Start from the second item
                if item is not None:
                    file.write(str(item) + "\n")  # Append each item with a newline character

def main():
    
    while True:
           
        # Check if the configuration file for Word documents exists
        folder_path = read_folder_path()

        if not folder_path:
            print("Configuration file for Word documents not found or folder path not configured.")
            folder_path = get_folder_path()

        # Check if the configuration file for Excel files exists
        excel_file_path = read_excel_file_path()

        if not excel_file_path:
            print("Configuration file for Excel files not found or Excel file path not configured.")
            excel_file_path = get_excel_file_path()

        if not os.path.exists(excel_file_path):
            print(f"Excel file '{excel_file_path}' not found.")
            return

        # Input: Provide a partial number to search for
        partial_number = input("Enter a partial number to search for: ")

        folder_dir = folder_path # Path to the folder with the Word documents

        # Search for folders with the same leading number as the partial number
        matching_folders = []
        for folder_name in os.listdir(folder_dir):
            if os.path.isdir(os.path.join(folder_dir, folder_name)) and folder_name.startswith(partial_number):
                matching_folders.append(folder_name)

        # Check if any matching folders were found
        if not matching_folders:
            print(f"No matching folders found for '{partial_number}' in Word documents.")
            return

        # Determine the directory of the executable (script or .exe)
        if getattr(sys, 'frozen', False):
        # The script is running as a compiled executable (.exe)
            script_dir = os.path.dirname(sys.executable)
        else:
        # The script is running as a regular Python script
            script_dir = os.path.dirname(__file__)

        # Iterate through the matching folders and extract data from the .docx files
        for folder_name in matching_folders:
            folder_path = os.path.join(folder_dir, folder_name)
            docx_files = [file for file in os.listdir(folder_path) if file.endswith(".docx") and file.startswith(folder_name)]

            if not docx_files:
                print(f"No .docx files with the same leading number found in folder '{folder_name}'.")
                continue

            docx_file = os.path.join(folder_path, docx_files[0])  # Use the first .docx file found

            word_data = extract_word_data(docx_file) # Extract data from the .docx file

            # Path for the output text file in the same directory as the script
            output_file_path = os.path.join(script_dir, f"{folder_name}.txt")

            # Write the collected HP numbers to the output text file
            with open(output_file_path, "w") as output_file:
                for hp_number in word_data:
                    output_file.write(hp_number + "\n")

            print(f"Data from {docx_files[0]} has been saved to {folder_name}.txt")
            
        # Input: Provide a 4-digit number to search for in the Excel sheet (xlsx file)
        search_value = partial_number #input("Enter a 4-digit number to search for in the Excel sheet: ")

        # Search the Excel sheet and extract data
        extracted_data = search_excel_and_extract_data(excel_file_path, search_value)

        if extracted_data is None:
            print(f"No data found for '{search_value}' in the Excel sheet.")
            return

        # Display the extracted data
        cell_value = extracted_data[0] if extracted_data is not None else None

        if cell_value is not None:
            print(f"Cell Value: {cell_value}")
        else:
            print(f"No data found for '{search_value}' in the Excel sheet.")

        # Create a text file with the cell value as the filename and add the data to the file
        text_file_path = os.path.join(script_dir, f"{cell_value}.txt")
        add_data_to_text_file(text_file_path, extracted_data)
                    
        # Start thread to read file and display gui
        if os.path.exists(output_file_path):
            with open(output_file_path, "r") as output_file:
                lines = output_file.readlines()
                if lines:
                    # Determine the required window height based on the number of lines
                    window_height = len(lines) * 50

                    # Create a separate thread to run the GUI
                    gui_thread_thread = threading.Thread(target=gui_thread, args=(lines,))
                    gui_thread_thread.start()

        # Ask for next or close
        confirmation = input("Do you want to exit the program? (y/n): ")
        if not (confirmation.lower() == "no" or confirmation.lower() == "n"):
            print("Exiting the program.")
            break
        elif (confirmation.lower() == "no" or confirmation.lower() == "n"):
            print("------------------------------------------------------------------------")
            

if __name__ == "__main__":
    main()
