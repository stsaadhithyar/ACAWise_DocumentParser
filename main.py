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

s3_file = 'S3 URI'
session = boto3.client('textract',
                      region_name='us-east-1',
                      aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
                      aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'))
textract_json = call_textract(input_document=s3_file, features = [Textract_Features.TABLES,Textract_Pretty_Print.FORMS],boto3_textract_client=session)
doc = Document(textract_json)

workbook = openpyxl.Workbook()

#required form field
required_data =  []
with open('Form_field.txt', 'r') as f:
    for line in f:
        required_data.append(line.strip())

for page_idx, page in enumerate(doc.pages, start=1):
    tables = page.tables
    part_I = {}
    main = {}
    if len(tables) == 2:
        table_0 = tables[0]
        table_1 = tables[1]

        for field in page.form.fields:
            part_I[str(field.key)] = str(field.value)
            
        # Getting the required Data
        for i in required_data:
            main[i] = part_I.get(i, "")

        row_list = [] 
        for row in table_1.rows:
            temp = []
            for cell in row.cells:
                temp.append(cell.text)
            row_list.append(temp)
        
        # Select the active sheet 
        sheet = workbook.create_sheet(title=f"Page_{page_idx}")

        # Write "Part I" to the first row in column  and merge cells
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(main) + 1)
        sheet.cell(row=1, column=1).value = "Part I"
        sheet.cell(row=1, column=1).font = Font(bold=True)

        # Write to the Excel file
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
        for row_idx, row_data in enumerate(row_list[1:], start=row):  # Exclude the first row
            for col_idx, cell_value in enumerate(row_data):
                sheet.cell(row=row_idx, column=col_idx + 1).value = cell_value
    # Second sheet structure            
    elif len(tables) == 1:
        table = tables[0]
        all_rows = []
        for i, row in enumerate(table.rows):
            if i <= 1:
                continue
            all_rows.append([cell.text for cell in row.cells])
        
        # Write data from the table to the sheet
        sheet = workbook.create_sheet(title=f"Page_{page_idx}")
        # Write column names
        column_names = [' ', 'First_Name', 'Middle_Initial', 'Last_Name', 'SSN or Other TIN', 'DOB (if SSN or Other TIN is not available)', 'all 12 months ', 'Jan ', 'Feb ', 'Mar ', 'Apr ', 'May ', 'June ', 'July ', 'Aug ', 'Sept ', 'Oct ', 'Nov ', 'Dec ']
        for col_idx, column_name in enumerate(column_names, start=1):
            sheet.cell(row=1, column=col_idx).value = column_name
        # Write data rows
        for row_idx, row_data in enumerate(all_rows, start=2):
            for col_idx, cell_value in enumerate(row_data, start=1):
                sheet.cell(row=row_idx, column=col_idx).value = cell_value

if len(workbook.sheetnames) > 1:
    workbook.remove(workbook.worksheets[0])

workbook.save('GMC1095c1_parsed.xlsx')
