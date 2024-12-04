from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from docx import Document

def copy_text_and_style(worksheet, cell_address, paragraph):
    """
    Copy text and basic style from a Word paragraph to an Excel cell.
    """
    cell = worksheet[cell_address]
    cell.value = paragraph.text

    # Apply basic formatting
    cell.font = Font(name='Calibri', size=11)
    if paragraph.style.font.bold:
        cell.font = Font(bold=True)
    if paragraph.style.font.italic:
        cell.font = Font(italic=True)
    cell.alignment = Alignment(horizontal='left', wrap_text=True)

def copy_table_to_excel(worksheet, start_row, table):
    """
    Copy a Word table to an Excel worksheet starting at the specified row.
    """
    current_row = start_row
    for row in table.rows:
        current_col = 1
        for cell in row.cells:
            excel_cell = worksheet.cell(row=current_row, column=current_col)
            excel_cell.value = cell.text
            excel_cell.alignment = Alignment(horizontal='left', wrap_text=True)
            current_col += 1
        current_row += 1
    return current_row

import os

def export_word_to_excel(word_path, excel_path):
    """
    Export content from a Word document to an Excel file.
    """
    # Load Word document
    doc = Document(word_path)

    # Create Excel workbook and worksheet
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "WordToExcel"

    row = 1

    # Loop through elements in the Word document
    for element in doc.element.body:
        if element.tag.endswith('p'):  # Paragraph
            try:
                # Safely access paragraph if it exists
                paragraph = doc.paragraphs[row - 1] if row - 1 < len(doc.paragraphs) else None
                if paragraph and paragraph.text.strip():  # Ensure paragraph exists and is not empty
                    cell_address = f"A{row}"
                    copy_text_and_style(worksheet, cell_address, paragraph)
                    row += 1
            except IndexError as e:
                print(f"Skipping invalid paragraph at index {row - 1}: {e}")
        elif element.tag.endswith('tbl'):  # Table
            # Get the table object
            try:
                table = next((t for t in doc.tables if t._element == element), None)
                if table:
                    row = copy_table_to_excel(worksheet, row, table)
            except Exception as e:
                print(f"Skipping invalid table: {e}")

    # Ensure the output directory exists
    output_dir = os.path.dirname(excel_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Save Excel file
    workbook.save(excel_path)

# Paths
word_file_path = "/Users/mutai/Desktop/KERICHO_COUNTY_FINANCE.docx"
excel_file_path = "output/WordToExcel.xlsx"

# Export the Word document to Excel
export_word_to_excel(word_file_path, excel_file_path)
