from sheet import Sheet
from commands import Command
from timer import Timer
import os

os.system('cls' if os.name == 'nt' else 'clear')

timer = Timer()
sheet = Sheet('1abQSainHd7j44v2Wq8EToCS_5v22rMekwJcUu19mjtE')

cmd = sheet.steps_tab.read_next_command()
while cmd is not None:
    Command(cmd).exec(sheet)
    cmd = sheet.steps_tab.read_next_command()

sheet.summarize()

sheet.flush()

print(f'\n✔ Done {timer.check()}')