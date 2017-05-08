#!python3
import sys
import json
import openpyxl, pprint
from pprint import pprint


nistControls = dict()

def parse():
    workbook = openpyxl.load_workbook('NIST Control Baseline.xlsx')

    sheet = workbook.get_sheet_by_name('Control Baseline')

    for row in range(3, 279):
        controlNum = rowValue = sheet['A' + str(row)].value

        if not '-' in str(controlNum):
            continue

        rowValue = sheet['G' + str(row)].value

        if rowValue == "Not Selected" or rowValue == None:
            continue

        if ',' in str(rowValue):
            split = rowValue.split(",")

            nistControls[controlNum] = '1'
            print (controlNum)

            for s in split:
                nistControls[controlNum + ' ('+str(s).rstrip()+')'] = '1'

                print (controlNum + ' ('+str(s).rstrip()+')')

        else:
            nistControls[controlNum] = '1'
            print (controlNum)


parse()
