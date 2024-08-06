from sheet import Sheet, Tab, StepsTab
import gspread

sheet = Sheet('1abQSainHd7j44v2Wq8EToCS_5v22rMekwJcUu19mjtE')
stepsTab = StepsTab(sheet)

stepsTab.readNextCommand()
stepsTab.readNextCommand()
stepsTab.readNextCommand()
stepsTab.readNextCommand()
stepsTab.readNextCommand()
