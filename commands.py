from typing import List, Tuple

from sheet import Sheet, Tab

from utils import period_index, ensure, parallel_calls
import gapi
import consts

from timer import Timer

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
            case 'group':
                cmd_group(sheet, self.args[1:])
            case 'summarize':
                cmd_summarize(sheet, self.args[1:])
            case _:
                print(f'? Command not recognized, ignored: {self.args[0].upper()}')


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
        raise(Exception(f'Tab "{args[0]} not found!'))
    sheet.tabs[args[0]].duplicate(clone = True,expand_periods = True)
    return

from pprint import pprint
def cmd_spawn(sheet: Sheet, args):
    if args[0] not in sheet.tabs:
        raise(Exception(f'Tab "{args[0]} not found!'))
    
    targets_raw = [t.strip() for t in str.split(args[1],',')]
    targets = [t.split(consts.FRIENDLY_NAME_DELIMITER) for t in targets_raw]
    source_tab = sheet.tabs[args[0]]
    for target in targets:
        sheet.raw_tab_count += 1
        gapi.duplicate_tab(source_tab.ref, f'-{target[0]}', sheet.raw_tab_count)
    sheet.flush()
    # now register the new ones
    timer = Timer()
    all_sheets = sheet.ref.worksheets()
    for s in all_sheets:
        for i, target in enumerate(targets):
            friendly_name = target[0] if len(target) < 2 else target[1]
            if s.title == f'-{target[0]}':
                new_tab = source_tab.register_duplicate(s)
                new_tab.set_friendly_name(friendly_name)

    print(f'✔ New tab registrations done. {timer.check()}')

    return

def parse_var(tab: Tab, arg: str) -> Tuple[Tuple[int, int], str]:
    if consts.COL_DELIMITER in arg:
        a, b = arg.split(consts.COL_DELIMITER)
        var = tab.get_var_rows(a)
        return [var, b]
    else:
        var = tab.get_var_rows(arg)
        return [var, 'p']

def cmd_map(sheet: Sheet, args):
    # map [source tab] [source var] [target tab] [target var]
    # map [source tab] [source var]:[source col] [target tab] [target var]:[target col]
    # map [source tab] [source var] [source col] [target tab] [target var] [target col]
    assertMinArgs(args, 4)
    s = sheet.get_tab(args[0])
    if len(args) == 4:
        t = sheet.get_tab(args[2])
        sv, scol = parse_var(s, args[1])
        tv, tcol = parse_var(t, args[3])
    else:
        # 6 arg syntax
        t = sheet.get_tab(args[3])
        sv, scol = parse_var(s, f'{args[1]}:{args[2]}')
        tv, tcol = parse_var(t, f'{args[4]}:{args[5]}')

    ensure(sv[1] == tv[1], f'Mismatch in multi-row variable heights: {args[0]}→{args[1]} and {args[2]}→{args[3]}.')

    # wip - row to multirow var

    # var may be multi-row

    ensure(s.get_col(scol) is not None, f'Source tab does not have column "{scol}".')
    ensure(t.get_col(tcol) is not None, f'Target tab does not have column "{tcol}".')

    # cases:
    # source is col, target is col
    if scol != 'p' and tcol != 'p':
        sources = s.get_var_col_refs(sv, scol)
        mappings = [f'=\'{s.ref.title}\'!{ref}' for ref in sources]
        t.update_cell(tv[0], t.get_col(tcol), mappings)
    
    # X source is periods, target is col --> error
    if scol == 'p' and tcol != 'p':
        ensure(False, f'Cannot map source var periods ({args[1]}) to a target var column ({args[3]}).')

    if tcol == 'p':
        if scol == 'p':
            # ✔ source is periods, target is periods
            sources = s.get_var_col_refs(sv, scol)
            mappings = [[f'=\'{s.ref.title}\'!{ref}' for ref in x] for x in sources]
        else:
            # ✔ source is col, target is periods
            sources = s.get_var_col_refs(sv, scol)
            # pick up the same source repeatedly
            mappings = [[f'=\'{s.ref.title}\'!{ref}'] * sheet.settings['periods'] for ref in sources]
        t.update_period_cells(tv[0], mappings)
        print(f'✔ Done.')

def cmd_trend(sheet: Sheet, args):
    assertMinArgs(args, 5)
    t = sheet.get_tab(args[0])
    tv, rows = t.get_var_rows(args[1])
    ensure(rows == 1, f'{args[1]} is a multi-row variable, trend cannot be performed on it.')
    startP, endP = get_period_range(args[2], True)
    periods = endP - startP
    startV = float(args[3])
    endV = float(args[4])
    method = args[5] if len(args) > 5 else 'linear'
    if method not in ['linear', 'expo']:
        method = 'linear'
        print('! Warning: Defaulting method to linear')
    incAdd: float = (endV - startV) / periods if method == 'linear' else 0
    incMul: float = pow(endV / startV, 1 / periods) if method == 'expo' else 1 

    count = endP - startP + 1
    cells: List[str] = [''] * count
    v: float = startV
    for i in range(len(cells)):
        cells[i] = v
        v = v * incMul + incAdd
    gapi.update_cells(t.ref, tv, t.get_pcol() + startP - 1, [cells])
    print(f'✔ Done.')
    return

def get_period_range(v: str, range_required: bool = False) -> Tuple[int, int]:
    if '-' in v:
        ps = v.split('-')
        return [period_index(ps[0]), period_index(ps[1])]
    else:
        ensure(range_required == False, f'{v} is not a multi-period range but this is required.')
        return [period_index(v), period_index(v)]

def cmd_bump(sheet: Sheet, args):
    assertMinArgs(args, 4)
    t = sheet.get_tab(args[0])
    tv, rows = t.get_var_rows(args[1])
    ensure(rows == 1, f'{args[1]} is a multi-row variable, bump cannot be performed on it.')

    startP, endP = get_period_range(args[2])
    count = endP - startP + 1

    v = float(args[3])
    cells: List[str] = [''] * count
    for i in range(len(cells)):
        cells[i] = v
    gapi.update_cells(t.ref, tv, t.get_pcol() + startP - 1, [cells])
    # t.ref.update_cells(cells, 'USER_ENTERED')
    print(f'✔ Done.')
    return

def cmd_group(sheet: Sheet, args):
    assertMinArgs(args, 2)
    label = args[0]
    tabs = [t.strip() for t in args[1].split(',')]
    sheet.add_tab_group(label, tabs)

# Utilities

def assertMinArgs(args, min):
    ensure(len(args) >= min, f'Not enough arguments, we need at least {min}.')
