# coding: utf-8

import os
from openpyxl import load_workbook

string_need = input("Enter the order you want to find: \n")

folderpath = os.path.split(os.path.realpath(__file__))[0]

filepath = folderpath + '\注文_1.xlsx'

# Load xlsx file
data = load_workbook(filepath, read_only=True, keep_links=False)

# Get sheet names in xlsx file
sheetnames = data.sheetnames

for m in range(0, len(sheetnames)):
    sheet_now = data[sheetnames[m]]

    for n in range(1, sheet_now.max_row+1):
        celldata = sheet_now.cell(n, 1).value   # Read data of each cell in first column

        if celldata is not None and string_need in celldata:
            print('row', n, 'in sheet :', sheetnames[m], ', Order Name:', celldata)

