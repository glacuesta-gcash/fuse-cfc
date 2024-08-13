from typing import List, Dict, Tuple

import re
import asyncio
from functools import partial

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

from utils import col_num_to_letter
from timer import Timer

import gapi

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
    async def parallel_calls(self, *partials):
        coroutines = [
            asyncio.create_task(
                asyncio.to_thread(partial)
            )
            for partial in partials
        ]
        tasks = await asyncio.gather(*coroutines)
        return tasks
    
    def __init__(self, sheetKey: str):
        print('⇨ Connecting to Sheet...')
        self.ref = client.open_by_key(sheetKey)
        print('✔ Connected.')
        self.tabs: Dict[str, Tab] = {}
        self.steps_tab: StepsTab = None
        self.summary_tab: SummaryTab = None
        all_sheets = self.ref.worksheets()
        self.raw_tab_count = len(all_sheets)

        self.settings = {
            'periods': 12,
            'summary-periods': 12,
            'summary-start': 1
        }
        self.summary_vars: List[Tuple[str, str]] = []

        # sweep first to remove all transient tabs, to avoid triggering duplicate tab error on summary spawn
        # also, pull values from all input and summary tabs
        ranges_to_read: List[str] = []
        for sheet in all_sheets:
            if sheet.title[0] == '-':
                # generated tab, for cleanup
                gapi.delete_tab(sheet)
                # self.ref.del_worksheet(t)
                self.raw_tab_count -= 1
                print(f'→ Tab "{sheet.title}" removed.')
            elif sheet.title == '_steps':
                ranges_to_read.append(sheet.title)
            elif sheet.title[0] == '_':
                # steps, summary, or input
                ranges_to_read.append(f'\'{sheet.title}\'!A1:1')
                ranges_to_read.append(f'\'{sheet.title}\'!A1:A')

        timer = Timer()
        print('→ Performing a parallel flush and read...')
        results = asyncio.run(self.parallel_calls(
            partial(gapi.flush_requests, self.ref), 
            partial(gapi.read_ranges, self.ref, ranges_to_read)
        ))
        print(f'✔ Done with parallel calls {timer.check()}')
        
        # cache tab headers
        raw_tab_vals = results[1]
        col_headers_cache: Dict[str,List[str]] = {}
        row_headers_cache: Dict[str,List[str]] = {}
        full_cache: Dict[str,List[List[str]]] = {}
        for key in raw_tab_vals:
            re_match = re.search(r'\'_(.*)\'\!', key)
            tab_name = re_match.group(1)
            if '_steps' in key:
                full_cache[tab_name] = raw_tab_vals[key]
            elif 'A1:A' in key:
                # header column
                col_headers_cache[tab_name] = [el[0] if el != [] else '' for el in raw_tab_vals[key]]
            elif 'A1:' in key:
                # header row
                row_headers_cache[tab_name] = raw_tab_vals[key][0] if len(raw_tab_vals[key]) > 0 else []

        for sheet in all_sheets:
            if sheet.title == '_steps':
                timer = Timer()
                print('✔ Capturing Steps tab...')
                self.steps_tab = StepsTab(sheet, full_cache)
                print(f'Done {timer.check()}')
            elif sheet.title == '_summary':
                print('✔ Capturing Summary tab...')
                self.summary_tab = self.register_summary_tab(
                    sheet, 
                    cached_row_headers=row_headers_cache, 
                    cached_col_headers=col_headers_cache
                    )
            elif sheet.title[0] == '_':
                self.register_tab(sheet, 
                                  cached_row_headers=row_headers_cache, 
                                  cached_col_headers=col_headers_cache
                                  )

        if self.steps_tab == None:
            raise('No Steps tab found!')
        if self.summary_tab == None:
            raise('No Summary tab found!')
        
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
    
    def flush(self):
        gapi.flush_requests(self.ref)

    def add_summary_var(self, var, method):
        self.summary_vars.append((var, method))

    def summarize(self):
        self.summary_tab.summarize()

class Tab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet, sheet: Sheet, copy_attributes_from: 'Tab' = None, cached_row_headers = [], cached_col_headers = []):
        timer = Timer()
        self.ref = worksheet
        self.sheet = sheet
        self.id = worksheet.id
        self.name = worksheet.title[1:]
        self.type = 'input' if worksheet.title[0] == '_' else 'dynamic'
        if copy_attributes_from == None:
            # cache var references
            colVars = cached_col_headers[self.name] if self.name in cached_col_headers else self.ref.col_values(1)
            self.vars = {str(value): row + 1 for row, value in enumerate(colVars) if value}
            # find p column
            row_vals = cached_row_headers[self.name] if self.name in cached_row_headers else self.ref.row_values(1)
            self.pcol = self.find_index_with_value(row_vals, 'p')
            self.gcol = self.find_index_with_value(row_vals, 'g')
        else:
            self.vars = copy_attributes_from.vars
            self.pcol = copy_attributes_from.pcol
            self.gcol = copy_attributes_from.gcol
        print(f'✔ Tab {self.ref.title} registered. {len(self.vars)} variable(s). Period column {"not " if self.pcol is None else ""}found. Period Group column {"not " if self.gcol is None else ""}found. {timer.check()}')

    def find_index_with_value(self, row_vals, val) -> int:
        try:
            index = row_vals.index(val) + 1
        except ValueError:
            index = None # no p column
        return index

    def duplicate(self, newTitle: str = '', clone: bool = False, expand_periods: bool = False) -> 'Tab':
        timer = Timer()
        if clone is False:
            if newTitle == '':
                raise(f'Title of new tab cannot be blank!')
            if newTitle in self.sheet.tabs:
                raise(f'Destination tab "{newTitle}" already exists!')
        else:
            newTitle = self.ref.title[1:]
            if newTitle in self.sheet.tabs and self.sheet.tabs[newTitle].type == 'dynamic':
                raise(f'Tab has already been cloned.')
        # duplicate_sheet is NOT cached because it immediately may be referenced
        newSheet = self.sheet.ref.duplicate_sheet(self.ref.id,new_sheet_name=f'-{newTitle}',insert_sheet_index=self.sheet.raw_tab_count)
        self.sheet.raw_tab_count += 1
        # newSheet.update_tab_color('ff0000')
        gapi.update_tab_color(newSheet, { 'red': 1, 'green': 0, 'blue': 0 })
        print(f'✔ Tab "-{newTitle}" created from {self.ref.title}. {timer.check()}')
        newTab = self.sheet.register_tab(newSheet, copyAttributesFrom=self)
        if expand_periods:
            newTab.expand_periods()
        return newTab
    
    def get_period_cells_for_row(self, row) -> List[gspread.cell.Cell]:
        cellList = self.ref.range(row, self.pcol, row, self.pcol + self.sheet.settings['periods'] - 1)
        return cellList
    
    def update_period_cells(self, row, cells: List[gspread.cell.Cell]):
        startRow = row
        startCol = self.pcol
        vals = [[c.value for c in cells]]
        gapi.update_cells(self.ref, startRow, startCol, vals)
    
    def expand_periods(self):
        if self.pcol is None:
            return
        # duplicate p column as needed
        gapi.duplicate_column(self.ref, self.pcol, self.sheet.settings['periods'] - 1)
        cells = self.get_period_cells_for_row(1)
        for i, cell in enumerate(cells):
            cell.value = f'P{i+1}'
        self.update_period_cells(1, cells)
        if self.gcol is not None and self.gcol > self.pcol:
            self.gcol += self.sheet.settings['periods']
        #self.ref.update_cells(cells)

class StepsTab:
    def __init__(self, worksheet: gspread.worksheet.Worksheet, cached_cells = None):
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
        print(f'  {len(self.steps)} steps found.')

    def read_next_command(self) -> List[str]:
        if self.cursor >= len(self.steps):
            return None
        args = self.steps[self.cursor]

        print(f'\n⇨ ({self.cursor+1}/{len(self.steps)}) {args[0].upper()} {str.join(", ",args[1:])}...')

        self.cursor += 1

        return args

class SummaryTab:
    def __init__(self, tab: Tab):
        self.tab = tab
        self.ref = tab.ref
        self.sheet = tab.sheet

        self.tab.type = 'summary'
        gapi.insert_column(self.ref, 1, 1)
        self.tab.pcol += 1

    def summarize(self):
        groups = self.period_group_count()

        print(self.sheet.summary_vars)

        groupRows: List[int] = []
        groupVals: List[List[str]] = []

        for sv in self.sheet.summary_vars:
            cellRefs: List[str] = []
            tabNames: List[str] = []

            # add rows
            baseRow = self.tab.vars[sv[0]]

            # map
            for t in self.sheet.tabs.values():
                if t.type == 'dynamic':
                    row = 0
                    if sv[0] in t.vars:
                        # print(f'{t.name} has {sv[0]}')
                        row += 1
                        cellValues = t.get_period_cells_for_row(t.vars[sv[0]])
                        cellRefs.append(f'=\'{t.ref.title}\'!{cellValues[0].address}')
                        tabNames.append(t.name)
                        # calculate and store group summaries
                        vs: List[str] = []
                        for n in range(0, groups):
                            if sv[1] == 'last':
                                v = f'={self.period_group_ref_for_last(baseRow + row, n)}'
                                pass
                            elif sv[1] in ['sum','average']:
                                # sum, average, or other worksheet function
                                v = f'={sv[1]}({self.period_group_range_ref_for_row(baseRow + row, n)})'
                            vs.append(v)
                        groupVals.append(vs)
                        groupRows.append(baseRow + row)

            gapi.duplicate_row(self.ref, baseRow, len(cellRefs))
            # add cell references
            cellValues = [[v] for v in cellRefs]
            tabNameValues = [[v] for v in tabNames]
            # self.summaryTab.ref.update_cells(cells, 'USER_ENTERED')
            gapi.update_cells(self.ref, baseRow + 1, self.tab.pcol, cellValues) # skip baseRow as that's the orig
            gapi.update_cells(self.ref, baseRow + 1, 2, tabNameValues) # skip baseRow as that's the orig
            # remove original row

        # extend periods
        self.tab.expand_periods()

        # add summary col groups
        if groups > 1:
            gapi.duplicate_column(self.ref, self.tab.gcol, groups - 1)
        group_labels = [[self.period_group_label(n) for n in range(0, groups)]]
        self.update_period_group_values_for_row(1, group_labels)

        # add group summaries
        for i in range(0,len(groupRows)):
            gapi.update_cells(self.ref, groupRows[i], self.tab.gcol, [groupVals[i]]) # put in [] to make it one row

    def period_group_count(self) -> int:
        return int((self.sheet.settings['periods'] - self.sheet.settings['summary-start'] + 1) / self.sheet.settings['summary-periods'])
    
    def period_group_label(self, groupIndex: int) -> int:
        return f'P{self.period_group_start(groupIndex)}-P{self.period_group_end(groupIndex)}'
    
    def period_group_start(self, groupIndex: int) -> int:
        return self.sheet.settings['summary-start'] + self.sheet.settings['summary-periods'] * groupIndex
    def period_group_end(self, groupIndex: int) -> int:
        return self.period_group_start(groupIndex + 1) - 1
    
    def period_group_range_ref_for_row(self, row: int, groupIndex: int) -> str:
        ref = f'{col_num_to_letter(self.tab.pcol + self.period_group_start(groupIndex)-1)}{row}:{col_num_to_letter(self.tab.pcol + self.period_group_end(groupIndex)-1)}{row}'
        return ref
    def period_group_ref_for_last(self, row: int, groupIndex: int) -> str:
        ref = f'{col_num_to_letter(self.tab.pcol + self.period_group_end(groupIndex)-1)}{row}'
        return ref
    
    def update_period_group_values_for_row(self, row: int, vals: List[any]):
        gapi.update_cells(self.ref, row, self.tab.gcol, vals)
