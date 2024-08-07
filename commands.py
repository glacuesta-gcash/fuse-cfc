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
            case 'trend':
                cmdTrend(sheet, self.args[1:])
            case 'bump':
                cmdBump(sheet, self.args[1:])
            case 'set':
                cmdSet(sheet, self.args[1:])

def cmdSet(sheet: Sheet, args):
    arg = args[0].lower()
    if arg in ['periods']:
        val = int(args[1])
        print(f'✔ Setting {arg} to {val}.')
        sheet.settings[arg] = val

def cmdBuild(sheet: Sheet, args):
    if args[0] not in sheet.tabs:
        raise(f'Tab "{args[0]} not found!')
    sheet.tabs[args[0]].duplicate(clone=True)
    return

def cmdSpawn(sheet: Sheet, args):
    if args[0] not in sheet.tabs:
        raise(f'Tab "{args[0]} not found!')
    sheet.tabs[args[0]].duplicate(newTitle=args[1])
    return

def cmdMap(sheet: Sheet, args):
    return

def cmdTrend(sheet: Sheet, args):
    if len(args) < 6:
        raise('Not enough arguments, we need at least 6.')
    if args[0] not in sheet.tabs:
        raise(f'Tab "{args[0]} not found!')
    t = sheet.tabs[args[0]]
    if args[1] not in t.vars:
        raise(f'Variable "{args[1]} not found!')
    startP = periodIndex(args[2])
    endP = periodIndex(args[3])
    periods = endP - startP
    startV = float(args[4])
    endV = float(args[5])
    method = args[6] if len(args) > 6 else 'linear'
    incAdd: float = (endV - startV) / periods if method == 'linear' else 0
    incMul: float = pow(endV / startV, 1 / periods) if method == 'expo' else 1 
    cells = t.getPeriodCellsForRow(t.vars[args[1]])
    v: float = startV
    for i, cell in enumerate(cells):
        if startP <= i + 1 <= endP:
            cell.value = v
            v = v * incMul + incAdd
    t.ref.update_cells(cells, 'USER_ENTERED')
    return

def cmdBump(sheet: Sheet, args):
    if len(args) < 4:
        raise('Not enough arguments, we need at least 4.')
    if args[0] not in sheet.tabs:
        raise(f'Tab "{args[0]} not found!')
    t = sheet.tabs[args[0]]
    if args[1] not in t.vars:
        raise(f'Variable "{args[1]} not found!')
    startP = periodIndex(args[2])
    v = float(args[3])
    cells = t.getPeriodCellsForRow(t.vars[args[1]])
    for i, cell in enumerate(cells):
        if startP <= i + 1:
            cell.value = v
    t.ref.update_cells(cells, 'USER_ENTERED')
    return

def periodIndex(s):
    return int(s[1:])