from sheet import Sheet
from commands import Command
from timer import Timer

timer = Timer()
sheet = Sheet('1abQSainHd7j44v2Wq8EToCS_5v22rMekwJcUu19mjtE')

cmd = sheet.stepsTab.readNextCommand()
while cmd is not None:
    Command(cmd).exec(sheet)
    cmd = sheet.stepsTab.readNextCommand()

sheet.summarize()

print(f'\n✔ Done (🕑 {timer.check()})')