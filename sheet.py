from typing import List, Dict, Tuple

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

from utils import periodIndex
from timer import Timer

# Define the scope
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

# Add credentials to the account
creds = Credentials.from_service_account_file('./credentials.json', scopes=scope)
client = gspread.authorize(creds)
# for raw API calls
service = build('sheets', 'v4', credentials=creds)

class Sheet:
    def __init__(self, sheetKey: str):
        print('⇨ Connecting to Sheet...')
        self.ref = client.open_by_key(sheetKey)
        print('✔ Connected.')
        self.tabs: Dict[str, Tab] = {}
        self.stepsTab: StepsTab = None
        self.summaryTab: Tab = None
        allSheets = self.ref.worksheets()
        self.rawTabCount = len(allSheets)

        self.settings = {
            'periods': 12,
            'summary-periods': 12,
            'summary-start': 'p1'
        }
        self.summaryVars: List[Tuple[str, str]] = []

        # sweep first to remove all transient tabs, to avoid triggering duplicate tab error on summary spawn
        for t in allSheets:
            if t.title[0] == '-':
                # generated tab, for cleanup
                deleteTab(t)
                # self.ref.del_worksheet(t)
                self.rawTabCount -= 1
                print(f'→ Tab "{t.title}" removed.')
        flushRequests(self.ref)

        for t in allSheets:
            if t.title == '_steps':
                print('✔ Capturing Steps tab...')
                self.stepsTab = StepsTab(t)
            elif t.title == '_summary':
                print('✔ Capturing Summary tab...')
                # todo: duplicate the summary tab to keep original pristine and to allow scenarios
                self.summaryTab = self.registerTab(t).duplicate('summary', clone=True, expandPeriods=False)
                self.summaryTab.type = 'summary'
                insertColumn(self.summaryTab.ref, 1, 1)
                self.summaryTab.pcol += 1
                print(self.summaryTab.vars)
            elif t.title[0] == '_':
                self.registerTab(t)

        if self.stepsTab == None:
            raise('No Steps tab found!')
        if self.summaryTab == None:
            raise('No Summary tab found!')

    def registerTab(self, sheet: gspread.worksheet.Worksheet, copyAttributesFrom: 'Tab' = None) -> 'Tab':
        newTab = Tab(sheet, self, copyAttributesFrom=copyAttributesFrom)
        self.tabs[sheet.title[1:]] = newTab # do not include prefix
        return newTab

    def summarize(self):
        print(self.summaryVars)
        for sv in self.summaryVars:
            cellRefs: List[str] = []
            tabNames: List[str] = []
            for t in self.tabs.values():
                if t.type == 'dynamic':
                    if sv[0] in t.vars:
                        print(f'{t.name} has {sv[0]}')
                        cellValues = t.getPeriodCellsForRow(t.vars[sv[0]])
                        cellRefs.append(f'=\'{t.ref.title}\'!{cellValues[0].address}')
                        tabNames.append(t.name)
            print(sv)
            # add rows
            baseRow = self.summaryTab.vars[sv[0]]
            duplicateRow(self.summaryTab.ref, baseRow, len(cellRefs))
            # add cell references
            cellValues = [[v] for v in cellRefs]
            tabNameValues = [[v] for v in tabNames]
            # self.summaryTab.ref.update_cells(cells, 'USER_ENTERED')
            updateCells(self.summaryTab.ref, baseRow + 1, self.summaryTab.pcol, cellValues) # skip baseRow as that's the orig
            updateCells(self.summaryTab.ref, baseRow + 1, 2, tabNameValues) # skip baseRow as that's the orig
            # remove original row
        # extend periods
        self.summaryTab.expandPeriods()
        # add summary col groups
        pass

    def flush(self):
        flushRequests(self.ref)

class Tab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet, sheet: Sheet, copyAttributesFrom: 'Tab' = None):
        timer = Timer()
        self.ref = worksheet
        self.sheet = sheet
        self.id = worksheet.id
        self.name = worksheet.title[1:]
        self.type = 'input' if worksheet.title[0] == '_' else 'dynamic'
        if copyAttributesFrom == None:
            # cache var references
            colVars = self.ref.col_values(1)
            self.vars = {str(value): row + 1 for row, value in enumerate(colVars) if value}
            # find p column
            rowVals = self.ref.row_values(1)
            try:
                self.pcol = rowVals.index('p') + 1
            except ValueError:
                self.pcol = None # no p column
        else:
            self.vars = copyAttributesFrom.vars
            self.pcol = copyAttributesFrom.pcol
        print(f'✔ Tab {self.ref.title} registered. {len(self.vars)} variable(s). Period column {"not " if self.pcol is None else ""}found. {timer.check()}')

    def duplicate(self, newTitle: str = '', clone: bool = False, expandPeriods: bool = False) -> 'Tab':
        timer = Timer()
        if clone is False:
            if newTitle == '':
                raise(f'Title of new tab cannot be blank!')
            if newTitle in self.sheet.tabs:
                raise(f'Destination tab "{newTitle}" already exists!')
        else:
            newTitle = self.ref.title[1:]
            if newTitle in self.sheet.tabs and self.sheet.tabs[newTitle].type == 'dynamic':
                raise(f'Tab has already been cloned.')
        # duplicate_sheet is NOT cached because it immediately may be referenced
        newSheet = self.sheet.ref.duplicate_sheet(self.ref.id,new_sheet_name=f'-{newTitle}',insert_sheet_index=self.sheet.rawTabCount)
        self.sheet.rawTabCount += 1
        newSheet.update_tab_color('ff0000')
        print(f'✔ Tab "-{newTitle}" created from {self.ref.title}. {timer.check()}')
        newTab = self.sheet.registerTab(newSheet, copyAttributesFrom=self)
        if expandPeriods:
            newTab.expandPeriods()
        return newTab
    
    def getPeriodCellsForRow(self, row) -> List[gspread.cell.Cell]:
        cellList = self.ref.range(row, self.pcol, row, self.pcol + self.sheet.settings['periods'] - 1)
        return cellList
    
    def updatePeriodCells(self, row, cells: List[gspread.cell.Cell]):
        startRow = row
        startCol = self.pcol
        vals = [[c.value for c in cells]]
        updateCells(self.ref, startRow, startCol, vals)
    
    def expandPeriods(self):
        if self.pcol is None:
            return
        # duplicate p column as needed
        duplicateColumn(self.ref, self.pcol, self.sheet.settings['periods'] - 1)
        cells = self.getPeriodCellsForRow(1)
        for i, cell in enumerate(cells):
            cell.value = f'P{i+1}'
        self.updatePeriodCells(1, cells)
        #self.ref.update_cells(cells)

class StepsTab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet):
        self.ref = worksheet
        self.steps = self.ref.get_all_values()
        self.steps = [step for step in self.steps if any(token != "" for token in step)]
        for step in self.steps:
            while step[-1] == "":
                step.pop()
        self.cursor = 0
        print(f'{len(self.steps)} steps found.')

    def readNextCommand(self) -> List[str]:
        if self.cursor >= len(self.steps):
            return None
        args = self.steps[self.cursor]

        print(f'\n⇨ ({self.cursor+1}/{len(self.steps)}) {args[0].upper()} {str.join(", ",args[1:])}...')

        self.cursor += 1

        return args
    
# raw calls
# request caching and flushing

def parseCellValue(value):
    if isinstance(value, str) and value.startswith('='):
        return {'formulaValue': value}
    elif isinstance(value, (int, float)):
        return {'numberValue': value}
    else:
        return {'stringValue': value}

def updateCells(sheet: gspread.worksheet.Worksheet, startRow, startCol, vals):
    # vals is rows downward, and then across; each row must be of same length
    rows = []
    for r in vals:
        rows.append({
            'values': [
                { 'userEnteredValue': parseCellValue(x) } for x in r
            ]
        })
    requests = [
        {
            'updateCells': {
                'rows': rows,
                'fields': 'userEnteredValue',
                'start': {
                    # GridCoordinate
                    'sheetId': sheet.id,
                    'rowIndex': startRow-1,
                    'columnIndex': startCol-1
                }
            }
        }
    ]
    queueRequests(requests)

def deleteTab(sheet: gspread.worksheet.Worksheet):
    requests = [
        {
            'deleteSheet': {
                'sheetId': sheet.id
            }
        }
    ]
    queueRequests(requests)

def insertColumn(sheet: gspread.worksheet.Worksheet, sourceCol: int, times: int = 1):
    requests = [
        # Request to insert a new column at index 1 (B)
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": sourceCol,
                    "endIndex": sourceCol + times
                }# ,
                # "inheritFromBefore": True
            }
        }
    ]
    queueRequests(requests)

def duplicateColumn(sheet: gspread.worksheet.Worksheet, sourceCol: int, times: int = 1):
    insertColumn(sheet, sourceCol, times)
    requests = [
        # Request to copy-paste the column with formatting
        {
            "copyPaste": {
                "source": {
                    "sheetId": sheet.id,
                    "startRowIndex": 0,
                    "endRowIndex": sheet.row_count,
                    "startColumnIndex": sourceCol - 1,
                    "endColumnIndex": sourceCol
                },
                "destination": {
                    "sheetId": sheet.id,
                    "startRowIndex": 0,
                    "endRowIndex": sheet.row_count,
                    "startColumnIndex": sourceCol,
                    "endColumnIndex": sourceCol + times
                },
                "pasteType": "PASTE_NORMAL"
            }
        }
    ]
    queueRequests(requests)

def duplicateRow(sheet: gspread.worksheet.Worksheet, sourceRow: int, times: int = 1):
    requests = [
        # Request to insert a new column at index 1 (B)
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "ROWS",
                    "startIndex": sourceRow,
                    "endIndex": sourceRow + times
                },
                "inheritFromBefore": True
            }
        },
        # Request to copy-paste the column with formatting
        {
            "copyPaste": {
                "source": {
                    "sheetId": sheet.id,
                    "startColumnIndex": 0,
                    "endColumnIndex": sheet.col_count,
                    "startRowIndex": sourceRow - 1,
                    "endRowIndex": sourceRow
                },
                "destination": {
                    "sheetId": sheet.id,
                    "startColumnIndex": 0,
                    "endColumnIndex": sheet.col_count,
                    "startRowIndex": sourceRow,
                    "endRowIndex": sourceRow + times
                },
                "pasteType": "PASTE_NORMAL"
            }
        }
    ]
    queueRequests(requests)

requestQueue: List[any] = []
def queueRequests(requests):
    global requestQueue

    requestQueue.extend(requests)

from pprint import pprint
def flushRequests(spreadsheet: gspread.spreadsheet.Spreadsheet):
    global requestQueue

    if len(requestQueue) == 0:
        print('No commands queued to flush.')
        return
    
    pprint(requestQueue)

    print(f'Executing {len(requestQueue)} queued command(s)...', end='')
    # Execute the requests
    body = {
        'requests': requestQueue
    }

    timer = Timer()
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet.id, body=body).execute()

    print(f'done {timer.check()}.')
    
    requestQueue = []