
"""
Основной исполняемый модуль

"""
from readready import readyReadFile
from readservicenote import serviceNoteReadFile
from settings import READY_FILE, SN_FILE

if __name__ == "__main__":
    readyDf = readyReadFile(READY_FILE)
    serviceNoteDf = serviceNoteReadFile(SN_FILE)
    print(serviceNoteDf)
    serviceNoteDf.to_excel("output1.xlsx")
# readyReadFile(READY_FILE)
# readyReadFile(READY_FILE).to_excel("output.xlsx")
# serviceNoteReadFile(SN_FILE).to_excel("output1.xlsx")
# serviceNoteReadFile(SN_FILE)