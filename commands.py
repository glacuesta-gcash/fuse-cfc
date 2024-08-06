from typing import List

from sheet import Sheet, Tab

class Settings:
    def __init__(self):
        self.periods = 12

settings = Settings()

class Command:
    def __init__(self, args: List[str]):
        self.args = args
    def exec(self, sheet: Sheet):
        match self.args[0].lower():
            case 'build':
                cmdBuild(sheet, self.args[1:])
            case 'spawn':
                cmdSpawn(sheet, self.args[1:])
            case 'map':
                cmdMap(sheet, self.args[1:])
            case 'grow':
                cmdGrow(sheet, self.args[1:])
            case 'periods':
                cmdSet(sheet, self.args[1:])

def cmdSet(sheet: Sheet, args):
    global settings
    match args[0].lower():
        case 'periods':
            settings.periods = int(args[1])
            print(f'✔ Setting PERIODS to {settings.periods}.')

def cmdBuild(sheet: Sheet, args):
    pass

def cmdSpawn(sheet: Sheet, args):
    if sheet.tabs[args[0]] == None:
        raise(f'Source tab "{args[0]}" not found!')
    if args[1] in sheet.tabs:
        raise(f'Destination tab "{args[1]}" already exists!')
    newSheet = sheet.ref.duplicate_sheet(sheet.tabs[args[0]].id,new_sheet_name=f'-{args[1]}',insert_sheet_index=sheet.rawTabCount)
    sheet.rawTabCount += 1
    newSheet.update_tab_color('ff0000')
    print(f'✔ Tab "{args[1]}" created.')
    sheet.registerTab(newSheet)
    # Now to extend periods
    return

def cmdMap(sheet: Sheet, args):
    return

def cmdGrow(sheet: Sheet, args):
    return

