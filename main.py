from sheet import Sheet
from commands import Command

sheet = Sheet('1abQSainHd7j44v2Wq8EToCS_5v22rMekwJcUu19mjtE')

cmd = sheet.stepsTab.readNextCommand()
while cmd is not None:
    Command(cmd).exec(sheet)
    cmd = sheet.stepsTab.readNextCommand()

print('✔ Done.')