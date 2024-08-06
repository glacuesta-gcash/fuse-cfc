from typing import List, Dict

import gspread

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Add credentials to the account
creds = gspread.service_account('./credentials.json')

class Sheet:
    def __init__(self, sheetKey: str):
        print('⇨ Connecting to Sheet...')
        self.ref = creds.open_by_key(sheetKey)
        print('✔ Connected.')
        self.tabs: Dict[str, Tab] = {}
        self.stepsTab = None
        allSheets = self.ref.worksheets()
        self.rawTabCount = len(allSheets)
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
    def registerTab(self, sheet: gspread.worksheet.Worksheet):
        self.tabs[sheet.title[1:]] = Tab(sheet) # do not include prefix

class Tab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet):
        self.ref = worksheet
        self.id = worksheet.id
        # cache var references
        colVars = self.ref.col_values(1)
        self.vars = {str(value): row + 1 for row, value in enumerate(colVars) if value}
        # find p column
        rowVals = self.ref.row_values(1)
        try:
            self.pcol = rowVals.index('p')
        except ValueError:
            self.pcol = None # no p column
        print(f'✔ Tab {self.ref.title} registered. {len(self.vars)} variable(s). Period column {"not " if self.pcol is None else ""}found.')
    def expandPeriods():
        pass

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
    