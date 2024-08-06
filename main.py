import gspread

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Add credentials to the account
creds = gspread.service_account('./credentials.json')

# Get the instance of the Spreadsheet
sheet = creds.open_by_key('1abQSainHd7j44v2Wq8EToCS_5v22rMekwJcUu19mjtE')

# Get the first sheet of the Spreadsheet
worksheet = sheet.get_worksheet(0)

# Get all values from the first row
rows = worksheet.row_values(1)
print(rows)
