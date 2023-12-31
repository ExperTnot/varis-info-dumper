import os
import re
import time
from docx import Document

CONFIG_FILE = "config.txt"

# Function to extract data from a Word document
def extract_word_data(docx_file):
    document = Document(docx_file)
    data = []
    
    for paragraph in document.paragraphs:
        text = paragraph.text
        occurrences = []
        start = 0
        
        while True:
            start = text.find("HP:", start)
            
            if start == -1:
                break
            
            end = start + 10  # 10 characters (HP:0000000)
            hp_number = text[start:end]
            occurrences.append(hp_number)
            start = end
        
        if occurrences:
            data.extend(occurrences)
    
    return data

# Function to open a file with a 1-second delay
def open_file_with_delay(file_path):
    time.sleep(1)
    os.system(f"start {file_path}")

# Function to get the folder path from the user and save it to the configuration file
def get_folder_path():
    folder_path = input("Enter the folder path where your .docx files are located: ")
    with open(CONFIG_FILE, "w") as config_file:
        config_file.write(folder_path)
    return folder_path

# Function to read the folder path from the configuration file
def read_folder_path():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as config_file:
            return config_file.read()
    else:
        return None

# Main function to collect and consolidate data
def main():
    # Check if the configuration file exists
    folder_path = read_folder_path()

    if not folder_path:
        print("Configuration file not found or folder path not configured.")
        folder_path = get_folder_path()

    # Input: Provide a partial number to search for
    partial_number = input("Enter a partial number to search for: ")

    folder_dir = folder_path

    matching_folders = []
    for folder_name in os.listdir(folder_dir):
        if os.path.isdir(os.path.join(folder_dir, folder_name)) and folder_name.startswith(partial_number):
            matching_folders.append(folder_name)

    if not matching_folders:
        print(f"No matching folders found for '{partial_number}'.")
        return

    script_dir = os.path.dirname(__file__)

    for folder_name in matching_folders:
        folder_path = os.path.join(folder_dir, folder_name)
        docx_files = [file for file in os.listdir(folder_path) if file.endswith(".docx") and file.startswith(folder_name)]

        if not docx_files:
            print(f"No .docx files with the same leading number found in folder '{folder_name}'.")
            continue

        docx_file = os.path.join(folder_path, docx_files[0])  # Use the first .docx file found

        word_data = extract_word_data(docx_file)

        # Path for the output text file in the same directory as the script
        output_file_path = os.path.join(script_dir, f"{folder_name}.txt")

        # Write the collected HP numbers to the output text file
        with open(output_file_path, "w") as output_file:
            for hp_number in word_data:
                output_file.write(hp_number + "\n")

        print(f"Data from {docx_files[0]} has been saved to {folder_name}.txt")

        # Open the generated .txt file with a 1-second delay
        open_file_with_delay(output_file_path)

if __name__ == "__main__":
    main()
