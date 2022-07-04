# Automate_excel_hide-delete

import pandas as pd
import openpyxl
import re 

#wb = openpyxl.load_workbook(r"Downloads\\XL_Flash_HOURLY_KPI_24_06_2022_20-05Hrs.xlsx", read_only = False , data_only= True )
sheet = wb.worksheets[0]

row_count = sheet.max_row
column_count = sheet.max_column

ws = wb.active
for row in range(7, 116, 4):
    for col in range(1, 35):
        #print(row, ws.cell(row=row, column =col).value)
        if str(ws.cell(row=row, column =col).value) == "#DIV/0!":
            ws.cell(row=row, column =col).value = "None"
            #print(row, ws.cell(row=row, column =col).value)
        else:
            pass

#wb.save(r"C:\Users\shadwive\automation\XL_Flash_edited.xlsx")

import pandas as pd
import openpyxl
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.formatting.rule import IconSet, FormatObject
from openpyxl.formatting import Rule
from openpyxl.styles import Font, PatternFill, Border
from openpyxl.styles.differential import DifferentialStyle
import re 

wb = openpyxl.load_workbook(r"Documents\\XL_Flash_HOURLY_KPI_24_06_2022_20-05Hrs.xlsx", read_only = False , data_only= True )
sheet = wb.worksheets[0]

row_count = sheet.max_row
column_count = sheet.max_column

ws = wb.active
for row in range(7, 116, 4):
    for col in range(1, 35):
        #print(row, ws.cell(row=row, column =col).value)
        if str(ws.cell(row=row, column =col).value) == "#DIV/0!":
            ws.cell(row=row, column =col).value = ""
            #print(row, ws.cell(row=row, column =col).value)
        else:
            pass

wb.save(r"Documents\\XL_Flash_HOURLY_KPI_24_06_2022_20-05Hrs.xlsx")
