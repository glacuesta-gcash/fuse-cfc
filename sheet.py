from typing import List

import gspread

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Add credentials to the account
creds = gspread.service_account('./credentials.json')

class Sheet:
    def __init__(self, sheetKey: str):
        self.ref = creds.open_by_key(sheetKey)

class Tab:
    def __init__(self, sheet: Sheet, tabName: str):
        self.ref = sheet.ref.worksheet(tabName)

class StepsTab(Tab):
    def __init__(self, sheet: Sheet):
        super().__init__(sheet, '_steps')
        self.rows = self.ref.get_all_values()
        self.cursor = 1
    def readNextCommand(self) -> List[str]:
        command = self.rows[self.cursor]
        while command[-1] == "":
            command.pop()
        self.cursor += 1

        print(str.join(' ', command))

        return command
    