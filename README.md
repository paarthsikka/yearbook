# Yearbook Automation Script

This Python script automates the creation of yearbook pages by processing student photos and quotes from an Excel file, generating a well-formatted Word document. This tool is designed to simplify and speed up the process of compiling student yearbooks.

## Features
- **Photo Processing:** Automatically crops and resizes student photos to fit a specified aspect ratio.
- **Quote Management:** Inserts student quotes beneath their photos, ensuring they adhere to character limits and content guidelines.
- **Batch Processing:** Handles large datasets efficiently, processing multiple entries in one run.
- **Error Handling:** Manages missing or invalid data gracefully, using default placeholders where necessary.

## Technologies Used
- **Python Libraries:** `pandas`, `python-docx`, `Pillow`, `requests`, `opencv`, `pillow_heif`
- **File Formats:** Processes images in various formats (JPEG, PNG, HEIC) and handles Excel files.

## Usage
1. Ensure you have Python installed on your system.
2. Install required libraries:
    ```bash
    pip install pandas python-docx pillow requests opencv-python pillow_heif
    ```
3. Place your Excel file (e.g., `fnfin.xlsx`) in the script directory.
4. Run the script:
    ```bash
    python yearbook_script.py
    ```
