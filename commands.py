from typing import List

from sheet import Sheet, Tab

from utils import period_index

class Command:
    def __init__(self, args: List[str]):
        self.args = args
    def exec(self, sheet: Sheet):
        match self.args[0].lower():
            case 'build':
                cmd_build(sheet, self.args[1:])
            case 'spawn':
                cmd_spawn(sheet, self.args[1:])
            case 'map':
                cmd_map(sheet, self.args[1:])
            case 'trend':
                cmd_trend(sheet, self.args[1:])
            case 'bump':
                cmd_bump(sheet, self.args[1:])
            case 'set':
                cmd_set(sheet, self.args[1:])
            case 'summarize':
                cmd_summarize(sheet, self.args[1:])

def cmd_summarize(sheet: Sheet, args):
    assertMinArgs(args, 2)
    if args[0] not in sheet.summary_tab.vars:
        print(f'Var {args[0]} does not exist in Summary tab!')
        return
    sheet.summary_vars.append((args[0], args[1]))

def cmd_set(sheet: Sheet, args):
    arg = args[0].lower()
    if arg in ['periods', 'summary-start', 'summary-periods']:
        val = int(period_index(args[1])) if arg == 'summary-start' else int(args[1])
        print(f'✔ Setting {arg} to {val}.')
        sheet.settings[arg] = val

def cmd_build(sheet: Sheet, args):
    if args[0] not in sheet.tabs:
        raise(f'Tab "{args[0]} not found!')
    sheet.tabs[args[0]].duplicate(clone=True,expand_periods=True)
    return

def cmd_spawn(sheet: Sheet, args):
    if args[0] not in sheet.tabs:
        raise(f'Tab "{args[0]} not found!')
    sheet.tabs[args[0]].duplicate(newTitle=args[1])
    return

def cmd_map(sheet: Sheet, args):
    assertMinArgs(args, 4)
    s = getTab(sheet, args[0])
    t = getTab(sheet, args[2])
    sv = getVar(s, args[1])
    tv = getVar(t, args[3])

    sources = s.get_period_cells_for_row(sv)
    cells = t.get_period_cells_for_row(tv)
    for i, cell in enumerate(cells):
        cell.value = f'=\'{s.ref.title}\'!{sources[i].address}'
    t.ref.update_cells(cells, 'USER_ENTERED')
    print(f'✔ Done.')
    return

def cmd_trend(sheet: Sheet, args):
    assertMinArgs(args, 6)
    t = getTab(sheet, args[0])
    tv = getVar(t, args[1])
    startP = period_index(args[2])
    endP = period_index(args[3])
    periods = endP - startP
    startV = float(args[4])
    endV = float(args[5])
    method = args[6] if len(args) > 6 else 'linear'
    incAdd: float = (endV - startV) / periods if method == 'linear' else 0
    incMul: float = pow(endV / startV, 1 / periods) if method == 'expo' else 1 
    cells = t.get_period_cells_for_row(tv)
    v: float = startV
    for i, cell in enumerate(cells):
        if startP <= i + 1 <= endP:
            cell.value = v
            v = v * incMul + incAdd
    # t.ref.update_cells(cells, 'USER_ENTERED')
    t.update_period_cells(tv,cells)
    print(f'✔ Done.')
    return

def cmd_bump(sheet: Sheet, args):
    assertMinArgs(args, 4)
    t = getTab(sheet, args[0])
    tv = getVar(t, args[1])
    startP = period_index(args[2])
    v = float(args[3])
    cells = t.get_period_cells_for_row(tv)
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