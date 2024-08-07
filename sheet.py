from typing import List, Dict

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

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
        self.stepsTab = None
        allSheets = self.ref.worksheets()
        self.rawTabCount = len(allSheets)

        self.settings = {
            'periods': 12
        }

        for t in allSheets:
            if t.title == '_steps':
                print('✔ Capturing Steps tab...')
                self.stepsTab = StepsTab(t)
            elif t.title[0] == '_':
                self.registerTab(t)
            elif t.title[0] == '-':
                # generated tab, for cleanup
                self.ref.del_worksheet(t)
                self.rawTabCount -= 1
                print(f'→ Tab "{t.title}" removed.')
        if self.stepsTab == None:
            raise('No Steps tab found!')
    def registerTab(self, sheet: gspread.worksheet.Worksheet) -> 'Tab':
        newTab = Tab(sheet, self)
        self.tabs[sheet.title[1:]] = newTab # do not include prefix
        return newTab

class Tab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet, sheet: Sheet):
        self.ref = worksheet
        self.sheet = sheet
        self.id = worksheet.id
        self.name = worksheet.title[1:]
        self.type = 'input' if worksheet.title[0] == '_' else 'dynamic'
        # cache var references
        colVars = self.ref.col_values(1)
        self.vars = {str(value): row + 1 for row, value in enumerate(colVars) if value}
        # find p column
        rowVals = self.ref.row_values(1)
        try:
            self.pcol = rowVals.index('p') + 1
        except ValueError:
            self.pcol = None # no p column
        print(f'✔ Tab {self.ref.title} registered. {len(self.vars)} variable(s). Period column {"not " if self.pcol is None else ""}found.')

    def duplicate(self, newTitle: str = '', clone: bool = False):
        if clone is False:
            if newTitle == '':
                raise(f'Title of new tab cannot be blank!')
            if newTitle in self.sheet.tabs:
                raise(f'Destination tab "{newTitle}" already exists!')
        else:
            newTitle = self.ref.title[1:]
            if newTitle in self.sheet.tabs and self.sheet.tabs[newTitle].type == 'dynamic':
                raise(f'Tab has already been cloned.')
        newSheet = self.sheet.ref.duplicate_sheet(self.ref.id,new_sheet_name=f'-{newTitle}',insert_sheet_index=self.sheet.rawTabCount)
        self.sheet.rawTabCount += 1
        newSheet.update_tab_color('ff0000')
        print(f'✔ Tab "-{newTitle}" created from {self.ref.title}.')
        newTab = self.sheet.registerTab(newSheet)
        newTab.expandPeriods()
    
    def getPeriodCellsForRow(self, row) -> List[gspread.cell.Cell]:
        cellList = self.ref.range(row, self.pcol, row, self.pcol + self.sheet.settings['periods'] - 1)
        return cellList
    
    def expandPeriods(self):
        if self.pcol is None:
            return
        # duplicate p column as needed
        duplicateColumn(self.sheet.ref, self.ref, self.pcol, self.sheet.settings['periods'])
        cells = self.getPeriodCellsForRow(1)
        for i, cell in enumerate(cells):
            cell.value = f'P{i+1}'
        self.ref.update_cells(cells)

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

def duplicateColumn(spreadsheet: gspread.spreadsheet.Spreadsheet, sheet: gspread.worksheet.Worksheet, sourceCol: int, times: int = 1):
    requests = [
        # Request to insert a new column at index 1 (B)
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": sourceCol,
                    "endIndex": sourceCol + times - 1
                },
                "inheritFromBefore": True
            }
        },
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
                    "endColumnIndex": sourceCol + times - 1
                },
                "pasteType": "PASTE_NORMAL"
            }
        }
    ]

    # Execute the requests
    body = {
        'requests': requests
    }

    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet.id, body=body).execute()
