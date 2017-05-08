#!python3
import sys
import json
import openpyxl, pprint
from pprint import pprint


nistControls = dict()

def parse():
    workbook = openpyxl.load_workbook('MARS-E 2 0 Detailed Delta Table Final Version 02-11-2016.xlsx')

    sheet = workbook.get_sheet_by_name('Security Controls')

    for row in range(9, 392):
        rowValue = sheet['B' + str(row)].value
        if rowValue == '' or rowValue == None or rowValue.startswith("Total"): 
            continue

        nistControls[rowValue] = '1'

parse()