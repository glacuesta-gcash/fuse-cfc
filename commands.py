from typing import List, Tuple

from sheet import Sheet, Tab

from utils import period_index, ensure
import gapi

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
    if args[0] not in sheet.summary_tab.tab.vars:
        print(f'Var {args[0]} does not exist in Summary tab!')
        return
    sheet.add_summary_var(args[0], args[1])

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

def parse_var(tab: Tab, arg: str) -> Tuple[int, str]:
    if ':' in arg:
        a, b = arg.split(':')
        row = tab.get_var_row(a)
        return [row, b]
    else:
        row = tab.get_var_row(arg)
        return [row, 'p']

def cmd_map(sheet: Sheet, args):
    # map [source tab] [source var] [target tab] [target var]
    # map [source tab] [source var]:[source col] [target tab] [target var]:[target col]
    assertMinArgs(args, 4)
    s = sheet.get_tab(args[0])
    t = sheet.get_tab(args[2])
    sv, scol = parse_var(s, args[1])
    tv, tcol = parse_var(t, args[3])

    # var may be multi-row

    # cases:
    # source is col, target is col
    # TO-DO
    
    # X source is periods, target is col --> error
    if scol == 'p' and tcol != 'p':
        ensure(False, f'Cannot map source var periods ({args[1]}) to a target var column ({args[3]}).')

    if tcol == 'p':
        cells = t.get_period_cells_for_row(tv)
        if scol == 'p':
            # ✔ source is periods, target is periods
            sources = s.get_period_cells_for_row(sv)
            for i, cell in enumerate(cells):
                cell.value = f'=\'{s.ref.title}\'!{sources[i].address}'
        else:
            # ✔ source is col, target is periods
            # pick up the same source repeatedly
            source = s.get_row_col_ref(sv, s.get_col(scol))
            for i, cell in enumerate(cells):
                cell.value = f'={source}'
            print(cells)
        t.update_period_cells(tv, cells)
        # t.ref.update_cells(cells, 'USER_ENTERED')
        print(f'✔ Done.')

def cmd_trend(sheet: Sheet, args):
    assertMinArgs(args, 6)
    t = sheet.get_tab(args[0])
    tv, rows = t.get_var_rows(args[1])
    ensure(rows == 1, f'{args[1]} is a multi-row variable, trend cannot be performed on it.')
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
    t = sheet.get_tab(args[0])
    tv, rows = t.get_var_rows(args[1])
    ensure(rows == 1, f'{args[1]} is a multi-row variable, bump cannot be performed on it.')
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
    ensure(len(args) >= min, f'Not enough arguments, we need at least {min}.')
