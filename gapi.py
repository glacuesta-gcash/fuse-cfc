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
    print(f'✔ {len(ranges)} range(s) read. {timer.check()}')
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

def update_tab_color(sheet: gspread.worksheet.Worksheet, color):
    requests = [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet.id,
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

request_queue: List[any] = []
def queue_requests(requests):
    global request_queue

    request_queue.extend(requests)

from pprint import pprint
def flush_requests(spreadsheet: gspread.spreadsheet.Spreadsheet):
    global request_queue
    global service

    if len(request_queue) == 0:
        print('No commands queued to flush.')
        return
    
    # pprint(request_queue)
    print(f'  [{str.join(', ', [list(req.keys())[0] for req in request_queue])}]')

    print(f'  Executing {len(request_queue)} queued command(s)...', end='')
    # Execute the requests
    body = {
        'requests': request_queue
    }

    timer = Timer()
    response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet.id, body=body).execute()

    print(f'done {timer.check()}.')
    
    request_queue = []
    