#!python3
from collections import OrderedDict
import openpyxl, pprint
import xml.etree.cElementTree as ET
import MarsE

ns = {'controls': 'http://scap.nist.gov/schema/sp800-53/feed/2.0',
      'control': 'http://scap.nist.gov/schema/sp800-53/2.0' }

columns = { 
    'Family' : 'A',
    '#' : 'B',
    'Title' : 'C',
    'Method' : 'D',
    'Test Objective' : 'E',
    'Results' : 'F',
    'Assessment' : 'G',
    'Recommendation' : 'H',
}

workbook = openpyxl.Workbook()

controlFamilies = [
        'AC', 
        'AT',
        'AU',
        'CA',
        'CM',
        'CP',
        'IA',
        'IR',
        'MA',
        'MP',
        'PE',
        'PL',
        'PS',
        'RA',
        'SA',
        'SC',
        'SI',
        'PM',
        ]
controlFamily = "Documents"
title = ""

sheet = workbook.get_sheet_by_name('Sheet')
sheet.title = controlFamily

controlObjectives = OrderedDict()
lastControlNumber = ''
currentRow = 2

tree = ET.ElementTree(file='800-53a-objectives.xml')
root = tree.getroot()

def quote(s):
    return '"' + s + '"'

def getColumn(column, row):
    global columns
    return columns[column] + str(row)

# Generate Assessment
for control in tree.findall("controls:control", ns):
    controlNumber = control.find('control:number', ns)
    controlFamily = controlNumber.text.split('-')[0]
    title = control.find('control:title', ns)

    if controlFamily not in controlFamilies:
        continue;

    if controlNumber.text not in MarsE.nistControls:
        continue;

    lastControlNumber = controlNumber.text
    lastControlNumber = lastControlNumber.split('-')[1]

    assessments = control.find('control:potential-assessments', ns)
    title = control.find('control:title', ns)

    if assessments == None:
        continue;

    for a in assessments:

        objects = a.findall('control:object', ns)
        method = a.attrib['method']

        for o in objects:
            print (controlNumber.text + " " + method + " " + o.text)
            sheet[getColumn('Family', currentRow)].value = controlFamily
            sheet[getColumn('#', currentRow)].value = controlNumber.text
            sheet[getColumn('Title', currentRow)].value = title.text
            sheet[getColumn('Method', currentRow)].value = method
            sheet[getColumn('Documents', currentRow)].value = o.text

            currentRow = currentRow + 1

    # Process control-enhancements
    enhancements = control.find('control:control-enhancements', ns)

    if enhancements == None:
        continue

    for enhancement in enhancements:
        controlNumber = enhancement.find('control:number', ns)
        controlFamily = controlNumber.text.split('-')[0]

        if controlFamily not in controlFamilies:
            continue;
        
        if controlNumber.text not in MarsE.nistControls:
            continue;

        assessments = control.find('control:potential-assessments', ns)

        for a in assessments:

            objects = a.findall('control:object', ns)
            method = a.attrib['method']

            for o in objects:
                print (controlNumber.text + " " + method + " " + o.text)
                sheet[getColumn('Family', currentRow)].value = controlFamily
                sheet[getColumn('#', currentRow)].value = controlNumber.text
                sheet[getColumn('Title', currentRow)].value = title.text
                sheet[getColumn('Method', currentRow)].value = method
                sheet[getColumn('Documents', currentRow)].value = o.text

            currentRow = currentRow + 1


sheet[getColumn('Family', 1)].value = 'Family'
sheet[getColumn('#', 1)].value = '#'
sheet[getColumn('Title', 1)].value = 'Title'
sheet[getColumn('Method', 1)].value = 'Method'
sheet[getColumn('Documents', 1)].value = 'Documents'

workbook.save('NistDocuments.xlsx')

