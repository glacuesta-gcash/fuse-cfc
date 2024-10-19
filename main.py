from sheet import Sheet
from commands import Command
from timer import Timer
import os

os.system('cls' if os.name == 'nt' else 'clear')

id_orig =       '1abQSainHd7j44v2Wq8EToCS_5v22rMekwJcUu19mjtE'
id_base =       '1s0Cnb5o2vbXAYZinCbsMEYx5vIS7JlHrxAsVB2cjqtU'
id_aggressive = '1MB_DtVpHV5wImG3_qUj6nwdwxS72zta-wD4D8URe2NY'

timer = Timer()
sheet = Sheet(id_aggressive)

cmd = sheet.steps_tab.read_next_command()
while cmd is not None:
    Command(cmd).exec(sheet)
    cmd = sheet.steps_tab.read_next_command()

sheet.summarize()

sheet.flush()

print(f'\n✔ Done {timer.check()}')