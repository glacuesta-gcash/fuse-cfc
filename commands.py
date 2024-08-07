from typing import List

from sheet import Sheet, Tab

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
    arg = args[0].lower()
    if arg in ['periods']:
        val = int(args[1])
        print(f'✔ Setting {arg} to {val}.')
        sheet.settings[arg] = val

def cmdBuild(sheet: Sheet, args):
    if sheet.tabs[args[0]] == None:
        raise(f'Source tab "{args[0]}" not found!')
    sheet.tabs[args[0]].duplicate(clone=True)
    return

def cmdSpawn(sheet: Sheet, args):
    if sheet.tabs[args[0]] == None:
        raise(f'Source tab "{args[0]}" not found!')
    sheet.tabs[args[0]].duplicate(newTitle=args[1])
    return

def cmdMap(sheet: Sheet, args):
    return

def cmdGrow(sheet: Sheet, args):
    return

