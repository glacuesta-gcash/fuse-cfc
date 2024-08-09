from typing import List

from sheet import Sheet, Tab

from utils import periodIndex

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
            case 'summarize':
                cmdSummarize(sheet, self.args[1:])

def cmdSummarize(sheet: Sheet, args):
    assertMinArgs(args, 2)
    if args[0] not in sheet.summaryTab.vars:
        print(f'Var {args[0]} does not exist in Summary tab!')
        return
    sheet.summaryVars.append((args[0], args[1]))

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
    assertMinArgs(args, 4)
    s = getTab(sheet, args[0])
    t = getTab(sheet, args[2])
    sv = getVar(s, args[1])
    tv = getVar(t, args[3])

    sources = s.getPeriodCellsForRow(sv)
    cells = t.getPeriodCellsForRow(tv)
    for i, cell in enumerate(cells):
        cell.value = f'=\'{s.ref.title}\'!{sources[i].address}'
    t.ref.update_cells(cells, 'USER_ENTERED')
    print(f'✔ Done.')
    return

def cmdTrend(sheet: Sheet, args):
    assertMinArgs(args, 6)
    t = getTab(sheet, args[0])
    tv = getVar(t, args[1])
    startP = periodIndex(args[2])
    endP = periodIndex(args[3])
    periods = endP - startP
    startV = float(args[4])
    endV = float(args[5])
    method = args[6] if len(args) > 6 else 'linear'
    incAdd: float = (endV - startV) / periods if method == 'linear' else 0
    incMul: float = pow(endV / startV, 1 / periods) if method == 'expo' else 1 
    cells = t.getPeriodCellsForRow(tv)
    v: float = startV
    for i, cell in enumerate(cells):
        if startP <= i + 1 <= endP:
            cell.value = v
            v = v * incMul + incAdd
    # t.ref.update_cells(cells, 'USER_ENTERED')
    t.updatePeriodCells(tv,cells)
    print(f'✔ Done.')
    return

def cmdBump(sheet: Sheet, args):
    assertMinArgs(args, 4)
    t = getTab(sheet, args[0])
    tv = getVar(t, args[1])
    startP = periodIndex(args[2])
    v = float(args[3])
    cells = t.getPeriodCellsForRow(tv)
    for i, cell in enumerate(cells):
        if startP <= i + 1:
            cell.value = v
    t.ref.update_cells(cells, 'USER_ENTERED')
    print(f'✔ Done.')
    return

# Utilities

def assertMinArgs(args, min):
    if len(args) < min:
        raise(f'Not enough arguments, we need at least {min}')

def getTab(sheet, tabName):
    if tabName not in sheet.tabs:
        raise(f'Tab "{tabName} not found!')
    return sheet.tabs[tabName]

def getVar(tab: Tab, varName):
    if varName not in tab.vars:
        raise(f'Variable "{varName}" not found in tab "{tab.name}!')
    return tab.vars[varName]