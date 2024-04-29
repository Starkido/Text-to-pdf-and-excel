# -*- coding: utf-8 -*-
"""
Created on Sat Mar 23 10:48:39 2024

@author: Starkiddev
"""

import pdfkit
from openpyxl import Workbook

# Configuration for wkhtmltopdf
config = pdfkit.configuration(wkhtmltopdf='C:\\Program Files\\wkhtmltopdf\\wkhtmltopdf.exe')

def text_to_pdf(text, output_file):
    # Convert text to PDF
    pdfkit.from_string(text, output_file, configuration=config)

def text_to_excel(text, output_file):
    # Split text into lines
    lines = text.split('\n')

    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active

    # Write each line to Excel
    for row_idx, line in enumerate(lines, start=1):
        ws.cell(row=row_idx, column=1, value=line)

    # Save Excel workbook
    wb.save(output_file)

if __name__ == "__main__":
    # Sample text input
    text_data = "text_data.txt"
    with open(text_data, "r") as file:
        text_display = file.read()
    
    #text_data = """
    #This is line 1.
    #This is line 2.
    #This is line 3.
    #"""

    # Convert text to PDF
    text_to_pdf(text_display, "newpdf.pdf")

    # Convert text to Excel
    text_to_excel(text_display, "newexcel.xlsx")