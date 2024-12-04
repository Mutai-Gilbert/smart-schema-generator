# Convert Word Documents to Excel

A Python script to convert the content of Microsoft Word documents (`.docx`) into Excel spreadsheets (`.xlsx`) while preserving text formatting and table structures.

## Features

- Extracts paragraphs and tables from Word documents.
- Maintains text formatting such as bold, italic, and font styles.
- Converts Word tables into corresponding Excel rows and columns.
- Supports creating new directories for output files if not present.

## Requirements

Ensure you have the following installed:

- Python 3.7 or higher
- Virtual environment with the required Python libraries

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/convert-word-documents-to-excel.git
   cd convert-word-documents-to-excel
2. Set up a virtual environment and activate it:
   ```bash
   python -m venv venv
   source venv/bin/activate   # On Windows: venv\Scripts\activate\

3. Install dependencies:
    ```bash
    pip install -r requirements.txt

Usage

1. Place your Word document (.docx) in the project directory.
2. Edit convert.py to specify the file paths:

    ```bash
    word_file_path = "path/to/your/document.docx"
    excel_file_path = "path/to/output/WordToExcel.xlsx"

3. Run the script:
    ```bash
    python convert.py
4. The converted Excel file will be saved in the specified output path.

Error Handling

Empty paragraphs or unsupported elements: The script skips invalid or empty content and logs the issue.
FileNotFoundError: Ensure the output directory exists or will be created automatically.
IndexError: This usually happens when accessing an element that doesn't exist. The script now safely skips these cases.
Contributions

Contributions are welcome! To contribute:

Fork this repository.
Create a new branch (git checkout -b feature-name).
Commit your changes (git commit -am 'Add some feature').
Push to the branch (git push origin feature-name).
Open a pull reques