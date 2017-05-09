#!python3
from collections import OrderedDict
import openpyxl, pprint
import xml.etree.cElementTree as ET
import MarsE

ns = {'controls': 'http://scap.nist.gov/schema/sp800-53/feed/2.0',
      'control': 'http://scap.nist.gov/schema/sp800-53/2.0'}

columns = {
    'Control Family': 'A',
    '#': 'B',
    'NIST ID': 'C',
    'Name': 'D',
    'Test Objective': 'E',
    'Results': 'F',
    'Assessment': 'G',
    'Recommendation': 'H',
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
controlFamily = "Controls"

sheet = workbook.get_sheet_by_name('Sheet')
sheet.title = controlFamily

controlObjectives = OrderedDict()
lastControlNumber = ''
currentRow = 2

tree = ET.ElementTree(file='800-53a-objectives.xml')
root = tree.getroot()


def quote(s):
    return '"' + s + '"'


def getColumn(column: object, row: object) -> object:
    global columns
    return columns[column] + str(row)


def addObjective(number, text):
    global controlObjectives, lastControlNumber, currentRow, sheet, controlFamily

    print(
        controlFamily + "," + lastControlNumber + "," + controlFamily + "-" + lastControlNumber + "," + number + "," + quote(
            text))
    controlObjectives[number] = text

    sheet[getColumn('Control Family', currentRow)].value = controlFamily
    if "(" in lastControlNumber:
        n = lastControlNumber.split("(")[0]
        sheet[getColumn('#', currentRow)].value = n
    else:
        sheet[getColumn('#', currentRow)].value = lastControlNumber

    sheet[getColumn('#', currentRow)].data_type = openpyxl.cell.Cell.TYPE_NUMERIC

    sheet[getColumn('NIST ID', currentRow)].value = controlFamily + "-" + lastControlNumber
    sheet[getColumn('Name', currentRow)].value = number
    sheet[getColumn('Test Objective', currentRow)].value = text

    if text == '':
        sheet[getColumn('Assessment', currentRow)].value = 'N/A'
    else:
        sheet[getColumn('Assessment', currentRow)].value = ''

    sheet[getColumn('Recommendation', currentRow)].value = ''

    currentRow = currentRow + 1


def processObjective(objective):
    o = objective.findall('control:objective', ns)

    if o == None:
        return

    for newObjective in o:
        number = newObjective.find('control:number', ns)
        decision = newObjective.find('control:decision', ns)

        if number == None and decision == None:
            continue

        if decision != None:
            addObjective(number.text, decision.text)
        else:
            addObjective(number.text, '')

        processObjective(newObjective)


# Generate Assessment
for control in tree.findall("controls:control", ns):
    controlNumber = control.find('control:number', ns)
    controlFamily = controlNumber.text.split('-')[0]

    if controlFamily not in controlFamilies:
        continue

    if controlNumber.text not in MarsE.nistControls:
        continue

    lastControlNumber = controlNumber.text
    lastControlNumber = lastControlNumber.split('-')[1]

    objective = control.find('control:objective', ns)
    decision = objective.find('control:decision', ns)
    objectives = objective.findall('control:objective', ns)

    addObjective(controlNumber.text, decision.text)

    for o in objectives:
        number = o.find('control:number', ns)
        decision = o.find('control:decision', ns)

        if number == None and decision == None:
            continue

        if decision != None:
            addObjective(number.text, decision.text)
        else:
            addObjective(number.text, '')

        processObjective(o)

    # Process control-enhancements
    enhancements = control.find('control:control-enhancements', ns)

    if enhancements == None:
        continue

    for enhancement in enhancements:
        controlNumber = enhancement.find('control:number', ns)
        controlFamily = controlNumber.text.split('-')[0]

        if controlFamily not in controlFamilies:
            continue

        if controlNumber.text not in MarsE.nistControls:
            continue

        lastControlNumber = controlNumber.text
        lastControlNumber = lastControlNumber.split('-')[1]

        objective = enhancement.find('control:objective', ns)
        decision = objective.find('control:decision', ns)
        objectives = objective.findall('control:objective', ns)
        addObjective(controlNumber.text, decision.text)

        for o in objectives:
            number = o.find('control:number', ns)
            decision = o.find('control:decision', ns)

            if number == None and decision == None:
                continue

            if decision != None:
                addObjective(number.text, decision.text)
            else:
                addObjective(number.text, '')

            processObjective(o)

sheet[getColumn('Control Family', 1)].value = 'Control Family'
sheet[getColumn('#', 1)].value = '#'
sheet[getColumn('NIST ID', 1)].value = 'NIST ID'
sheet[getColumn('Name', 1)].value = 'Name'
sheet[getColumn('Test Objective', 1)].value = 'Test Objective'
sheet[getColumn('Results', 1)].value = 'Results'
sheet[getColumn('Assessment', 1)].value = 'Assessment'
sheet[getColumn('Recommendation', 1)].value = 'Recommendation'

workbook.save('Controls.xlsx')
