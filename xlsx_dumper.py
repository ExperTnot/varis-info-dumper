import os
import openpyxl

CONFIG_FILE = "config_xlsx.txt"

# Function to search for a 4-digit number in an Excel sheet and extract data from a specific column (Column C -> Column I)
def search_excel_and_extract_data(excel_file, search_value):
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    extracted_data = None

    # Iterate through rows in Column C
    for row in sheet.iter_rows(min_row=2, values_only=True):
        cell_value = str(row[2])  # Column C is the third column (index 2)

        if search_value in cell_value:
            # Extract data from Column I (if available)
            column_i_value = None
            if len(row) > 8:  # Check if there are enough columns in the current row
                column_i_value = row[8]  # Column I (index 8)

            extracted_data = (cell_value, column_i_value)
            break  # Stop iterating after the first match

    wb.close()
    return extracted_data

# Function to get the Excel file path from the user and save it to the configuration file
def get_excel_file_path():
    excel_file_path = input("Enter the path to the Excel sheet (xlsx file): ")
    with open(CONFIG_FILE, "w") as config_file:
        config_file.write(excel_file_path)
    return excel_file_path

# Function to read the Excel file path from the configuration file
def read_excel_file_path():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as config_file:
            return config_file.read()
    else:
        return None

# Function to add data to a text file    
def add_data_to_text_file(file_path, data):
     with open(file_path, "a") as file:
        if data is not None:
            file.write(data + "\n")  # Append data with a newline character

def main():
    # Check if the configuration file exists
    excel_file_path = read_excel_file_path()

    if not excel_file_path:
        print("Configuration file not found or Excel file path not configured.")
        excel_file_path = get_excel_file_path()

    if not os.path.exists(excel_file_path):
        print(f"Excel file '{excel_file_path}' not found.")
        return

    # Input: Provide a 4-digit number to search for in the Excel sheet (xlsx file)
    search_value = input("Enter a 4-digit number to search for in the Excel sheet: ")

    # Search the Excel sheet and extract data
    extracted_data = search_excel_and_extract_data(excel_file_path, search_value)

    if extracted_data is None:
        print(f"No data found for '{search_value}' in the Excel sheet.")
        return

    # Display the extracted data
    cell_value, data = extracted_data
    print(f"Cell Value: {cell_value}")
    if data is not None:
        print(f"Data: {data}\n")
    else:
        print("No data found in Column I\n")

    # Determine the directory of the script
    script_directory = os.path.dirname(__file__)

    # Create a text file with the cell value as the filename and add the data to the file
    text_file_path = os.path.join(script_directory, f"{cell_value}.txt")
    add_data_to_text_file(text_file_path, data)

if __name__ == "__main__":
    main()
