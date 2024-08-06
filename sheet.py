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
        allTabs = self.ref.worksheets()
        self.rawTabCount = len(allTabs)
        for t in allTabs:
            if t.title == '_steps':
                print('✔ Capturing Steps tab...')
                self.stepsTab = StepsTab(t)
            elif t.title[0] == '_':
                print(f'✔ Capturing tab "{t.title[1:]}"...')
                self.tabs[t.title[1:]] = Tab(t) # do not include _
            elif t.title[0] == '-':
                # generated tab, for cleanup
                print(f'→ Removing transient tab "{t.title}"...')
                self.ref.del_worksheet(t)
                self.rawTabCount -= 1
        if self.stepsTab == None:
            raise('No Steps tab found!')

class Tab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet):
        self.ref = worksheet
        self.id = worksheet.id

class StepsTab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet):
        self.ref = worksheet
        self.rows = self.ref.get_all_values()
        self.cursor = 1
    def readNextCommand(self) -> List[str]:
        if self.cursor >= len(self.rows):
            return None
        command = self.rows[self.cursor]
        while command[-1] == "":
            command.pop()
        self.cursor += 1

        return command
    