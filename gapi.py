from typing import List, Tuple, Callable

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

from timer import Timer

service: any

def set_service(svc):
    global service
    service = svc

# raw calls
# request caching and flushing

def parse_cell_value(value):
    if isinstance(value, str) and value.startswith('='):
        return {'formulaValue': value}
    elif isinstance(value, (int, float)):
        return {'numberValue': value}
    else:
        return {'stringValue': value}

def read_ranges(spreadsheet: gspread.spreadsheet.Spreadsheet, ranges: List[str]):
    timer = Timer()
    result = service.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet.id,
        ranges=ranges
    ).execute()
    response = {}
    for value_range in result['valueRanges']:
        # this assumes first row and first col are non-blank, otherwise it will lead to parsing issues downstream
        if 'values' in value_range:
            values = value_range['values']
            if value_range['majorDimension'] == 'COLUMNS':
                values = list(zip(*values))
                values = [list(row) for row in values]
        else:
            values = []
        response[value_range['range']] = values
    print(f'  ✔ {len(ranges)} range(s) read. {timer.check()}')
    return response

def update_cells(sheet: gspread.worksheet.Worksheet, startRow, startCol, vals):
    # vals is rows downward, and then across; each row must be of same length
    rows = []
    for r in vals:
        rows.append({
            'values': [
                { 'userEnteredValue': parse_cell_value(x) } for x in r
            ]
        })
    requests = [
        {
            'updateCells': {
                'rows': rows,
                'fields': 'userEnteredValue',
                'start': {
                    # GridCoordinate
                    'sheetId': sheet.id,
                    'rowIndex': startRow-1,
                    'columnIndex': startCol-1
                }
            }
        }
    ]
    queue_requests(requests)

def delete_tab(sheet: gspread.worksheet.Worksheet):
    requests = [
        {
            'deleteSheet': {
                'sheetId': sheet.id
            }
        }
    ]
    queue_requests(requests)

def update_tab_color(sheet: gspread.worksheet.Worksheet | str, color):
    requests = [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet if isinstance(sheet, str) else sheet.id,
                    # "title": title, # In this case, I think that this might not be required to be used.
                    "tabColor": color
                },
                "fields": "tabColor"
            }
        }
    ]
    queue_requests(requests)

def insert_column(sheet: gspread.worksheet.Worksheet, sourceCol: int, times: int = 1):
    requests = [
        # Request to insert a new column at index 1 (B)
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": sourceCol,
                    "endIndex": sourceCol + times
                }# ,
                # "inheritFromBefore": True
            }
        }
    ]
    queue_requests(requests)

def group_columns(sheet: gspread.worksheet.Worksheet, startCol: int, endCol: int):
    requests = [
        {
            "addDimensionGroup": {
                "range": {
                    "dimension": "COLUMNS",
                    "sheetId": sheet.id,
                    "startIndex": startCol,
                    "endIndex": endCol
                }
            }
        },
        {
            "updateDimensionGroup": {
                "dimensionGroup": {
                    "range": {
                        "dimension": "COLUMNS",
                        "sheetId": sheet.id,
                        "startIndex": startCol,
                        "endIndex": endCol
                    },
                    "depth": 1,
                    "collapsed": True
                },
                "fields": "*"
            }
        }
    ]
    queue_requests(requests)

def group_rows(sheet: gspread.worksheet.Worksheet, startRow: int, endRow: int, collapse: bool = True):
    requests = [
        {
            "addDimensionGroup": {
                "range": {
                    "dimension": "ROWS",
                    "sheetId": sheet.id,
                    "startIndex": startRow,
                    "endIndex": endRow
                }
            }
        }
    ]
    if collapse:
        requests.append(
            {
                "updateDimensionGroup": {
                    "dimensionGroup": {
                        "range": {
                            "dimension": "ROWS",
                            "sheetId": sheet.id,
                            "startIndex": startRow,
                            "endIndex": endRow
                        },
                        "depth": 1,
                        "collapsed": True
                    },
                    "fields": "*"
                }
            }
        )
    queue_requests(requests)

def duplicate_column(sheet: gspread.worksheet.Worksheet, sourceCol: int, times: int = 1):
    insert_column(sheet, sourceCol, times)
    requests = [
        # Request to copy-paste the column with formatting
        {
            "copyPaste": {
                "source": {
                    "sheetId": sheet.id,
                    "startRowIndex": 0,
                    "endRowIndex": sheet.row_count,
                    "startColumnIndex": sourceCol - 1,
                    "endColumnIndex": sourceCol
                },
                "destination": {
                    "sheetId": sheet.id,
                    "startRowIndex": 0,
                    "endRowIndex": sheet.row_count,
                    "startColumnIndex": sourceCol,
                    "endColumnIndex": sourceCol + times
                },
                "pasteType": "PASTE_NORMAL"
            }
        }
    ]
    queue_requests(requests)

def duplicate_row(sheet: gspread.worksheet.Worksheet, sourceRow: int, times: int = 1):
    if times == 0:
        return
    requests = [
        # Request to insert a new column at index 1 (B)
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "ROWS",
                    "startIndex": sourceRow,
                    "endIndex": sourceRow + times
                },
                "inheritFromBefore": True
            }
        },
        # Request to copy-paste the column with formatting
        {
            "copyPaste": {
                "source": {
                    "sheetId": sheet.id,
                    "startColumnIndex": 0,
                    "endColumnIndex": sheet.col_count,
                    "startRowIndex": sourceRow - 1,
                    "endRowIndex": sourceRow
                },
                "destination": {
                    "sheetId": sheet.id,
                    "startColumnIndex": 0,
                    "endColumnIndex": sheet.col_count,
                    "startRowIndex": sourceRow,
                    "endRowIndex": sourceRow + times
                },
                "pasteType": "PASTE_NORMAL"
            }
        }
    ]
    queue_requests(requests)

def duplicate_tab(sheet: gspread.worksheet.Worksheet, new_sheet_name: str, index: int, after: Callable = None):
    requests = [
        {
            'duplicateSheet': {
                'sourceSheetId': sheet.id,
                'insertSheetIndex': index,
                'newSheetName': new_sheet_name
            }
        }
    ]
    queue_requests(requests, [after])

request_queue: List[any] = []
callback_queue: List[Callable] = []
def queue_requests(requests, callbacks: List[Callable] = None):
    global request_queue
    global callback_queue

    if callbacks == None:
        callbacks = [None] * len(requests)

    request_queue.extend(requests)
    callback_queue.extend(callbacks)

def flush_requests(spreadsheet: gspread.spreadsheet.Spreadsheet):
    global request_queue
    global callback_queue

    global service

    if len(request_queue) == 0:
        print('No commands queued to flush.')
        return
    
    print(f'→ Executing {len(request_queue)} queued command(s)...')
    if False: # set to True for verbose output
        print(f'  {[list(req.keys())[0] for req in request_queue]}')
    # Execute the requests
    body = {
        'requests': request_queue
    }

    timer = Timer()
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet.id, body=body).execute()

    for i, reply in enumerate(response['replies']):
        if callback_queue[i] is not None:
            callback_queue[i](reply)

    print(f'✔ ...done executing. {timer.check()}')
    
    request_queue = []
    callback_queue = []
    