import os
import re
from docx import Document

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

# Main function to collect and consolidate data
def main():
    # Input: Provide a partial number to search for
    partial_number = input("Enter a partial number to search for: ")

    folder_dir = r"D:\Code\varis_info\files"

    matching_folders = []
    for folder_name in os.listdir(folder_dir):
        if os.path.isdir(os.path.join(folder_dir, folder_name)) and partial_number in folder_name:
            matching_folders.append(folder_name)

    if not matching_folders:
        print(f"No matching folders found for '{partial_number}'.")
        return

    for folder_name in matching_folders:
        folder_path = os.path.join(folder_dir, folder_name)
        docx_files = [file for file in os.listdir(folder_path) if file.endswith(".docx")]

        if not docx_files:
            print(f"No .docx files found in folder '{folder_name}'.")
            continue

        docx_file = os.path.join(folder_path, docx_files[0])  # Use the first .docx file found

        word_data = extract_word_data(docx_file)

        # Path for the output text file in the same folder as the Word file
        output_file_path = os.path.join(folder_path, f"{folder_name}.txt")

        # Write the collected HP numbers to the output text file
        with open(output_file_path, "w") as output_file:
            for hp_number in word_data:
                output_file.write(hp_number + "\n")

        print(f"Data from {docx_files[0]} has been saved to {folder_name}.txt")

if __name__ == "__main__":
    main()
