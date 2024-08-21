from typing import List, Dict, Tuple

import re
import asyncio
from functools import partial

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

from utils import col_num_to_letter, ensure, parallel_calls, row_col_to_cell_ref
from timer import Timer
import gapi
import consts

# Define the scope
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

# Add credentials to the account
creds = Credentials.from_service_account_file('./credentials.json', scopes=scope)
client = gspread.authorize(creds)

# for raw API calls
gapi.set_service(build('sheets', 'v4', credentials=creds))

class Sheet:
    def __init__(self, sheetKey: str):
        timer = Timer()
        print('⇨ Connecting to Sheet...', end='')
        self.ref = client.open_by_key(sheetKey)
        print(f'connected. {timer.check()}')
        self.tabs: Dict[str, Tab] = {}
        self.steps_tab: StepsTab = None
        self.summary_tab: SummaryTab = None
        timer = Timer()
        all_sheets = self.ref.worksheets()
        print(f'  Sheets loaded. {timer.check()}')
        self.raw_tab_count = len(all_sheets)

        self.settings = {
            'periods': 12,
            'summary-periods': 12,
            'summary-start': 1
        }
        self.summary_vars: List[Tuple[str, str]] = []
        self.summary_tab_order: List[Tab] = []
        self.tab_groups: List[str] = []

        # sweep first to remove all transient tabs, to avoid triggering duplicate tab error on summary spawn
        # also, pull values from all input and summary tabs
        ranges_to_read: List[str] = []
        for sheet in all_sheets:
            if sheet.title[0] == consts.TAB_PREFIX_DYNAMIC:
                # generated tab, for cleanup
                gapi.delete_tab(sheet)
                # self.ref.del_worksheet(t)
                self.raw_tab_count -= 1
                print(f'→ Tab "{sheet.title}" removed.')
            elif sheet.title == consts.TAB_TITLE_STEPS:
                ranges_to_read.append(sheet.title)
            elif sheet.title[0] == consts.TAB_PREFIX_INPUT:
                # steps, summary, or input
                ranges_to_read.append(f'\'{sheet.title}\'!A1:1')
                ranges_to_read.append(f'\'{sheet.title}\'!A1:A')

        timer = Timer()
        print('→ Performing a parallel flush and read...')
        results = asyncio.run(parallel_calls(
            partial(gapi.flush_requests, self.ref), 
            partial(gapi.read_ranges, self.ref, ranges_to_read)
        ))
        print(f'  Done with parallel calls {timer.check()}')
        
        # cache tab headers
        raw_tab_vals = results[1]
        col_headers_cache: Dict[str,List[str]] = {}
        row_headers_cache: Dict[str,List[str]] = {}
        full_cache: Dict[str,List[List[str]]] = {}
        for key in raw_tab_vals:
            re_match = re.search(fr'\'{consts.TAB_PREFIX_INPUT}(.*)\'\!', key)
            tab_name = re_match.group(1)
            if consts.TAB_TITLE_STEPS in key:
                full_cache[tab_name] = raw_tab_vals[key]
            # elif 'A1:A' in key:
            elif re.search(r'A1:A\d+$', key) is not None:
                # header column
                col_headers_cache[tab_name] = [el[0] if el != [] else '' for el in raw_tab_vals[key]]
            # elif 'A1:' in key:
            elif re.search(r'A1:\w+1$', key) is not None:
                # header row
                row_headers_cache[tab_name] = raw_tab_vals[key][0] if len(raw_tab_vals[key]) > 0 else []

        for sheet in all_sheets:
            if sheet.title == consts.TAB_TITLE_STEPS:
                print('✔ Capturing Steps tab...')
                self.steps_tab = StepsTab(sheet, full_cache)
            elif sheet.title == consts.TAB_TITLE_SUMMARY:
                print('✔ Capturing Summary tab...')
                self.summary_tab = self.register_summary_tab(
                    sheet, 
                    cached_row_headers=row_headers_cache, 
                    cached_col_headers=col_headers_cache
                    )
            elif sheet.title[0] == consts.TAB_PREFIX_INPUT:
                self.register_tab(sheet, 
                                  cached_row_headers=row_headers_cache, 
                                  cached_col_headers=col_headers_cache
                                  )

        if self.steps_tab == None:
            raise(Exception('No Steps tab found!'))
        if self.summary_tab == None:
            raise(Exception('No Summary tab found!'))
        
    def register_summary_tab(self, sheet: gspread.worksheet.Worksheet, copyAttributesFrom: 'Tab' = None, cached_row_headers = [], cached_col_headers = []) -> 'SummaryTab':
        newTab = SummaryTab(
            self.register_tab(
                sheet, 
                cached_row_headers=cached_row_headers, 
                cached_col_headers=cached_col_headers
                ).duplicate('summary', clone=True, expand_periods=False)
            )
        return newTab

    def register_tab(self, sheet: gspread.worksheet.Worksheet, copyAttributesFrom: 'Tab' = None, cached_row_headers = [], cached_col_headers = []) -> 'Tab':
        newTab = Tab(
            sheet, 
            self, 
            copyAttributesFrom, 
            cached_row_headers, 
            cached_col_headers
            )
        self.tabs[sheet.title[1:]] = newTab # do not include prefix
        return newTab
    
    def get_tab(self, tab_name: str) -> 'Tab':
        ensure(tab_name in self.tabs, f'Tab "{tab_name}" not found!')
        return self.tabs[tab_name]

    def flush(self):
        gapi.flush_requests(self.ref)

    def add_summary_var(self, var, method):
        self.summary_vars.append((var, method))

    def summarize(self):
        # finish out the summary tab order with all other dynamic tabs in their natural order
        for tab in self.tabs.values():
            if tab not in self.summary_tab_order and tab.type == 'dynamic':
                self.summary_tab_order.append(tab)

        self.summary_tab.summarize()

    def add_tab_group(self, label: str, subtabs: List[str]):
        if label in self.tab_groups:
            raise(Exception(f'Tab group {label} already defined!'))
        self.tab_groups.append(label)
        for t in subtabs:
            tab = self.get_tab(t)
            if tab.group is not None:
                raise(Exception(f'Tab {t} already belongs to group {self.tabs[t].group}!'))
            tab.group = label
            self.summary_tab_order.append(tab)

class Tab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet, sheet: Sheet, copy_attributes_from: 'Tab' = None, cached_row_headers = [], cached_col_headers = []):
        timer = Timer()
        self.ref = worksheet
        self.sheet = sheet
        self.id = worksheet.id
        self.name = worksheet.title[1:]
        self.friendly_name = self.name
        self.group = None
        self.type = 'input' if worksheet.title[0] == consts.TAB_PREFIX_INPUT else 'dynamic'

        self.vars: Dict[str, Tuple[int, int]] = {} # label -> row, count
        self.cols: Dict[str, int] = {} # label -> col

        if copy_attributes_from == None:
            col_vars = cached_col_headers[self.name] if self.name in cached_col_headers else self.ref.col_values(1)
            # cache var references
            temp = {str(value): row + 1 for row, value in enumerate(col_vars) if value}
            for t in temp:
                if consts.VAR_HEIGHT_DELIMITER in t:
                    var, rows = t.split(consts.VAR_HEIGHT_DELIMITER)
                    self.vars[var] = [temp[t], int(rows)]
                else:
                    self.vars[t] = [temp[t], 1]
            # find p column
            row_vals = cached_row_headers[self.name] if self.name in cached_row_headers else self.ref.row_values(1)
            self.cols = {str(value): col + 1 for col, value in enumerate(row_vals) if value}
            self.pcol = self.get_pcol()
            self.gcol = self.get_gcol()
        else:
            self.vars = copy_attributes_from.vars
            self.cols = copy_attributes_from.cols
            self.pcol = copy_attributes_from.pcol
            self.gcol = copy_attributes_from.gcol
        print(f'✔ Tab {self.ref.title} registered. {len(self.vars)} variable(s). Period column {"not " if self.pcol is None else ""}found. {timer.check()}')

    def get_pcol(self) -> int:
        return None if 'p' not in self.cols else self.cols['p']
    def get_gcol(self) -> int:
        return None if 'g' not in self.cols else self.cols['g']
    
    def get_var_row(self, var: str) -> int:
        ensure(var in self.vars, f'Variable "{var}" not found in tab "{self.name}"!')
        return self.vars[var][0]
    def get_var_rows(self, var: str) -> Tuple[int, int]:
        ensure(var in self.vars, f'Variable "{var}" not found in tab "{self.name}"!')
        return self.vars[var]
    
    def nudge_var_row(self, var: str, delta: int):
        ensure(var in self.vars, f'Variable "{var}" not found in tab "{self.name}"!')
        self.vars[var][0] += delta

    def get_col(self, label: str) -> int:
        ensure(label in self.cols, f'Column "{label}" not found in tab "{self.name}"!')
        return self.cols[label]
    def nudge_col(self, label, delta):
        if label in self.cols:
            base = self.cols[label]
            for i in self.cols:
                if self.cols[i] >= base:
                    self.cols[i] += delta
    def nudge_pcol(self, delta):
        self.nudge_col('p', delta)
        self.pcol = self.get_col('p')
    def nudge_gcol(self, delta):
        self.nudge_col('g', delta)
        self.gcol = self.get_col('g')

    def get_row_col_ref(self, row, col, idx = 0) -> str:
        return f'\'{self.ref.title}\'!{col_num_to_letter(col)}{row + idx}'
    def get_var_col_refs(self, var: Tuple[int, int], col: str) -> List[str] | List[List[str]]:
        """May return a vertical stack if variable is multi-row"""
        baseRow = var[0]
        count = var[1]
        baseCol = self.get_col(col)
        if col == 'p':
            result = [[f'{col_num_to_letter(baseCol + x)}{baseRow + y}' for x in range(0, self.sheet.settings['periods'])] for y in range(0, count)]
        else:
            result = [f'{col_num_to_letter(baseCol)}{baseRow + y}' for y in range(0, count)]
        return result
    
    def duplicate(self, newTitle: str = '', clone: bool = False, expand_periods: bool = False) -> 'Tab':
        timer = Timer()
        if clone is False:
            if newTitle == '':
                raise(Exception(f'Title of new tab cannot be blank!'))
            if newTitle in self.sheet.tabs:
                raise(Exception(f'Destination tab "{newTitle}" already exists!'))
        else:
            newTitle = self.ref.title[1:]
            if newTitle in self.sheet.tabs and self.sheet.tabs[newTitle].type == 'dynamic':
                raise(Exception(f'Tab has already been cloned.'))
        # duplicate_sheet is NOT cached because it immediately may be referenced
        newSheet = self.sheet.ref.duplicate_sheet(self.ref.id,new_sheet_name=f'{consts.TAB_PREFIX_DYNAMIC}{newTitle}',insert_sheet_index=self.sheet.raw_tab_count)
        self.sheet.raw_tab_count += 1
        # newSheet.update_tab_color('ff0000')
        gapi.update_tab_color(newSheet, { 'red': 1, 'green': 0, 'blue': 0 })
        print(f'✔ Tab "-{newTitle}" created from {self.ref.title}. {timer.check()}')
        newTab = self.sheet.register_tab(newSheet, copyAttributesFrom=self)
        if expand_periods:
            newTab.expand_periods()
        return newTab
    
    def register_duplicate(self, new_sheet):
        gapi.update_tab_color(new_sheet, { 'red': 1, 'green': 0, 'blue': 0 })
        newTab = self.sheet.register_tab(new_sheet, copyAttributesFrom=self)
        newTab.expand_periods()
        return newTab
    
    def set_friendly_name(self, new_friendly_name: str):
        self.friendly_name = new_friendly_name

    def get_period_cells_for_row(self, row) -> List[gspread.cell.Cell]:
        cellList = self.ref.range(row, self.get_pcol(), row, self.get_pcol() + self.sheet.settings['periods'] - 1)
        return cellList
    
    def update_period_cells(self, row: int, vals: List[str | List[str]]):
        """If vals is a List of Lists then it's rows x cols, otherwise a single row."""
        startRow = row
        startCol = self.get_pcol()
        gapi.update_cells(self.ref, startRow, startCol, vals if isinstance(vals[0],List) else [vals])

    def update_cell(self, row: int, col: int, vals: str | List[str]):
        """Can accept a vertical stack, by passing List[str] to val."""
        gapi.update_cells(self.ref, row, col, [[c] for c in vals] if isinstance(vals, List) else [[vals]])
    
    def expand_periods(self):
        if self.get_pcol() is None:
            return
        # duplicate p column as needed
        gapi.duplicate_column(self.ref, self.get_pcol(), self.sheet.settings['periods'] - 1)
        cells: List[str] = [''] * self.sheet.settings['periods']
        for i in range(len(cells)):
            cells[i] = f'P{i+1}'
        self.update_period_cells(1, cells)
        if self.get_gcol() is not None and self.get_gcol() > self.get_pcol():
            self.nudge_gcol(self.sheet.settings['periods'] - 1)
        #self.ref.update_cells(cells)

class StepsTab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet, cached_cells = None):
        timer = Timer()
        self.ref = worksheet
        if cached_cells is not None and 'steps' in cached_cells:
            self.steps = cached_cells['steps']
        else:
            self.steps = self.ref.get_all_values()
        self.steps = [step for step in self.steps if any(token != "" for token in step)]
        for step in self.steps:
            while step[-1] == "":
                step.pop()
        self.cursor = 0
        print(f'  {len(self.steps)} steps found. {timer.check()}')

    def read_next_command(self) -> List[str]:
        if self.cursor >= len(self.steps):
            return None
        args = self.steps[self.cursor]

        print(f'\n⇨ ({self.cursor+1}/{len(self.steps)}) {args[0].upper()} {str.join(", ",args[1:])}...')

        self.cursor += 1

        return args
from pprint import pprint
class SummaryTab:
    def __init__(self, tab: Tab):
        self.tab = tab
        self.ref = tab.ref
        self.sheet = tab.sheet

        self.tab.type = 'summary'
        gapi.insert_column(self.ref, 1, 1)
        self.tab.nudge_pcol(1)

    def summarize(self):
        timer = Timer()
        print(f'\n→ Performing summary...',end='')

        pgroups = self.period_group_count()

        pgroupRows: List[int] = []
        pgroupVals: List[List[str]] = []

        # sort summary_vars according to position on summary tab, to avoid outdated cell refs in cell updates
        # lowest first
        self.sheet.summary_vars.sort(key=lambda sv: self.tab.get_var_row(sv[0]))

        # prepare tab groupings via mappings
        # sort tabs by tab groupings as declared

        for sv in self.sheet.summary_vars:
            cellRefs: List[str] = []
            tabNames: List[str] = []

            # add rows
            baseRow = self.tab.get_var_row(sv[0])

            # track groups with values
            groups_touched: Dict[str,List[int]] = {} # list is start, end row
            cur_group = None

            # walk each tab that has this variable
            row = 0
            for t in self.sheet.summary_tab_order:
                if t.type == 'dynamic':
                    if sv[0] in t.vars:
                        # yup this tab has this var

                        if t.group != cur_group:
                            # change group so tie the prev one off
                            if cur_group is not None:
                                groups_touched[cur_group][1] = baseRow + row

                        row += 1

                        if t.group is not None and t.group not in groups_touched.keys():
                            # new group
                            # add spacer to house subtotal for group
                            cellRefs.append('')
                            tabNames.append(t.group)
                            groups_touched[t.group] = [baseRow + row + 1, 0] # start collecting the next row, not this
                            row += 1
                            cur_group = t.group

                        # this is expensive. made it not need it
                        # cellValues = t.get_period_cells_for_row(t.get_var_row(sv[0]))
                        cellRefs.append(f'=\'{t.ref.title}\'!{row_col_to_cell_ref(t.get_var_row(sv[0]), t.get_pcol())}')
                        # indented if part of group
                        tabNames.append((consts.GROUP_INDENT if t.group is not None else "") + t.friendly_name)

            rows = row

            # wrap up group tracking
            if cur_group is not None:
                groups_touched[cur_group][1] = baseRow + row

            # do period groups for all rows added
            if rows > 0:
                for i in range(0, rows + 1):
                    vs: List[str] = []
                    for n in range(0, pgroups):
                        if sv[1] == 'last':
                            v = f'={self.period_group_ref_for_last(baseRow + i, n)}'
                            pass
                        elif sv[1] in ['sum','average']:
                            # sum, average, or other worksheet function
                            v = f'={sv[1]}({self.period_group_range_ref_for_row(baseRow + i, n)})'
                        vs.append(v)

                    # cached group values for later
                    pgroupVals.append(vs)
                    pgroupRows.append(baseRow + i)

            gapi.duplicate_row(self.ref, baseRow, rows)
            for k in self.tab.vars:
                if self.tab.get_var_row(k) > baseRow:
                    self.tab.nudge_var_row(k, len(cellRefs))

            # add cell references
            cellValues = [[v] for v in cellRefs]
            tabNameValues = [[v] for v in tabNames]
            # self.summaryTab.ref.update_cells(cells, 'USER_ENTERED')
            gapi.update_cells(self.ref, baseRow + 1, self.tab.get_pcol(), cellValues) # skip baseRow as that's the orig
            gapi.update_cells(self.ref, baseRow + 1, 2, tabNameValues) # skip baseRow as that's the orig

            # subtotals
            if rows > 0:
                # put sum in baseRow
                col_letter = col_num_to_letter(self.tab.get_pcol())
                v = f'=subtotal(9,{col_letter}{baseRow + 1}:{col_letter}{baseRow + rows})'
                self.tab.update_cell(baseRow, self.tab.get_pcol(), v)

                # now let's do the groups
                for g in groups_touched.values():
                    # periods
                    v = f'=subtotal(9,{col_letter}{g[0]}:{col_letter}{g[1]})'
                    self.tab.update_cell(g[0]-1, self.tab.get_pcol(), v)

                    # collapse group rows
                    gapi.group_rows(self.tab.ref, g[0]-1, g[1])

                gapi.group_rows(self.tab.ref, baseRow, baseRow + rows, collapse = False)

        # extend periods
        # this will then also capture and copy-paste the cell refs for the period cells (but doesn't work for period groups)
        self.tab.expand_periods()

        # collapse period cols
        gapi.group_columns(self.tab.ref, self.tab.get_pcol() - 1, self.tab.get_pcol() + self.sheet.settings['periods'] - 1)

        # add summary period group columns
        if pgroups > 1:
            gapi.duplicate_column(self.ref, self.tab.get_gcol(), pgroups - 1)
        pgroup_labels = [[self.period_group_label(n) for n in range(0, pgroups)]]
        self.update_period_group_values_for_row(1, pgroup_labels)

        # add period group summaries
        for i in range(0,len(pgroupRows)):
            gapi.update_cells(self.ref, pgroupRows[i], self.tab.get_gcol(), [pgroupVals[i]]) # put in [] to make it one row

        print(f'done. {timer.check()}\n')

    def period_group_count(self) -> int:
        return int((self.sheet.settings['periods'] - self.sheet.settings['summary-start'] + 1) / self.sheet.settings['summary-periods'])
    
    def period_group_label(self, groupIndex: int) -> int:
        return f'P{self.period_group_start(groupIndex)}-P{self.period_group_end(groupIndex)}'
    
    def period_group_start(self, groupIndex: int) -> int:
        return self.sheet.settings['summary-start'] + self.sheet.settings['summary-periods'] * groupIndex
    def period_group_end(self, groupIndex: int) -> int:
        return self.period_group_start(groupIndex + 1) - 1
    
    def period_group_range_ref_for_row(self, row: int, groupIndex: int) -> str:
        ref = f'{col_num_to_letter(self.tab.get_pcol() + self.period_group_start(groupIndex)-1)}{row}:{col_num_to_letter(self.tab.get_pcol() + self.period_group_end(groupIndex)-1)}{row}'
        return ref
    def period_group_ref_for_last(self, row: int, groupIndex: int) -> str:
        ref = f'{col_num_to_letter(self.tab.get_pcol() + self.period_group_end(groupIndex)-1)}{row}'
        return ref
    
    def update_period_group_values_for_row(self, row: int, vals: List[any]):
        gapi.update_cells(self.ref, row, self.tab.get_gcol(), vals)
