from typing import List

from sheet import Sheet, Tab

class Command:
    def __init__(self, args: List[str]):
        self.args = args
    def exec(self, sheet: Sheet):
        match self.args[0].lower():
            case 'spawn':
                cmdSpawn(sheet, self.args[1:])
            case 'map':
                cmdMap(sheet, self.args[1:])
            case 'grow':
                cmdGrow(sheet, self.args[1:])

def cmdSpawn(sheet: Sheet, args):
    print(f'\n⇨ Performing SPAWN with args {str.join(',',args)}...')
    if sheet.tabs[args[0]] == None:
        raise(f'Source tab "{args[0]}" not found!')
    if args[1] in sheet.tabs:
        raise(f'Destination tab "{args[1]}" already exists!')
    newSheet = sheet.ref.duplicate_sheet(sheet.tabs[args[0]].id,new_sheet_name=f'-{args[1]}',insert_sheet_index=sheet.rawTabCount)
    sheet.rawTabCount += 1
    newSheet.update_tab_color('ff0000')
    sheet.tabs[args[1]] = Tab(newSheet)
    print(f'✔ Tab "{args[1]}" created.')
    return

def cmdMap(sheet: Sheet, args):
    print(f'\n⇨ Performing MAP with args {str.join(',',args)}...')
    return

def cmdGrow(sheet: Sheet, args):
    print(f'\n⇨ Performing GROW with args {str.join(',',args)}...')
    return

