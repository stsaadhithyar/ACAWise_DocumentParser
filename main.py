import boto3
import json
import pandas as pd
from textractcaller import call_textract, Textract_Features
from textractprettyprinter.t_pretty_print import Pretty_Print_Table_Format, Textract_Pretty_Print, get_string, get_tables_string
from trp import Document
from trp.trp2 import TDocument, TDocumentSchema
from trp.t_pipeline import order_blocks_by_geo
from IPython.display import display
import openpyxl
from openpyxl.styles import Font
import os
from dotenv import load_dotenv

load_dotenv()

s3_file = 's3://textract-testing-sts004/GMC1095c1_update1.pdf'
session = boto3.client('textract',
                      region_name='us-east-1',
                      aws_access_key_id=os.getenv('ACCESS_KEY_ID'),
                      aws_secret_access_key=os.getenv('SECRET_ACCESS_KEY'))
textract_json = call_textract(input_document=s3_file, features = [Textract_Features.TABLES,Textract_Pretty_Print.FORMS],boto3_textract_client=session)

required_data =  []

with open('/home/sts852-aadhithyar/Documents/ACA/Main/ACA_Main/Form_field.txt', 'r') as f:
    for line in f:
        required_data.append(line.strip())

doc = Document(textract_json)
import openpyxl
from openpyxl.styles import Font

# Create a new Excel workbook
workbook = openpyxl.Workbook()
global row
row = 1

for page_idx, page in enumerate(doc.pages, start=1):
    tables = page.tables
    #part_1 dictionary has all the data form the form
    part_1 = {}
    #main dictionary contains only the required data which is needed
    main = {}
    if len(tables) == 2:
        table_0 = tables[0]
        table_1 = tables[1]

        # Data stored in dictionary
        for field in page.form.fields:
            part_1[str(field.key)] = str(field.value)
            
        # Getting the required Data
        for i in required_data:
            main[i] = part_1.get(i, "")

        l = []  # Initialize list outside the loop to store rows
        for row in table_1.rows:
            # Initialize an empty list to store the content of cells in the row
            temp = []
            for cell in row.cells:
                # Append the content of each cell to the temporary list
                temp.append(cell.text)
            # Append the content of the row to the main list
            l.append(temp)
        
        # Select the active sheet (create new sheet for each page)
        sheet = workbook.create_sheet(title=f"Page_{page_idx}")

        # Write "Part I" to the first row in column A and merge cells
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(main) + 1)
        sheet.cell(row=1, column=1).value = "Part I"
        sheet.cell(row=1, column=1).font = Font(bold=True)

        # Write the keys and values from the dictionary to the Excel file
        row = 3
        for key, value in main.items():
            sheet.cell(row=row, column=1).value = key
            sheet.cell(row=row, column=2).value = value
            row += 1

        row += 1
        # Write "Part II" in the next row
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(main) + 1)
        sheet.cell(row=row, column=1).value = "Part II"
        sheet.cell(row=row, column=1).font = Font(bold=True)
        row += 2

        # Write the table data to the Excel file
        for row_idx, row_data in enumerate(l[1:], start=row):  # Exclude the first row
            for col_idx, cell_value in enumerate(row_data):
                sheet.cell(row=row_idx, column=col_idx + 1).value = cell_value
            row += 1
        
    elif len(tables) == 1:
        row += 2
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(main) + 1)
        sheet.cell(row=row, column=1).value = "Part III"
        sheet.cell(row=row, column=1).font = Font(bold=True)
        row += 2
        table = tables[0]
        all_rows = []
        for i, rows in enumerate(table.rows):
            if i <= 1:
                continue
            all_rows.append([cell.text for cell in rows.cells])
        # Write data from the table to the sheet
        #sheet = workbook.create_sheet(title=f"Page_{page_idx}")
        # Write column names
        column_names = [' ', 'First_Name', 'Middle_Initial', 'Last_Name', 'SSN or Other TIN', 'DOB (if SSN or Other TIN is not available)', 'all 12 months ', 'Jan ', 'Feb ', 'Mar ', 'Apr ', 'May ', 'June ', 'July ', 'Aug ', 'Sept ', 'Oct ', 'Nov ', 'Dec ']
        for col_idx, column_name in enumerate(column_names, start=1):
            sheet.cell(row=row, column=col_idx).value = column_name
        row += 1
        # Write data rows
        for row_idx, row_data in enumerate(all_rows, start=2):
            for col_idx, cell_value in enumerate(row_data, start=1):
                sheet.cell(row=row, column=col_idx).value = cell_value
            row += 1

# Save the Excel workbook with a specified filename
workbook.save('/home/sts852-aadhithyar/Documents/ACA/Main/ACA_Main/Test/Output.xlsx')