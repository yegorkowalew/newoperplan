
"""
Основной исполняемый модуль

"""
from readready import readyReadFile
from readservicenote import serviceNoteReadFile
from settings import READY_FILE, SN_FILE

if __name__ == "__main__":
    print(readyReadFile(READY_FILE))
    print(serviceNoteReadFile(SN_FILE))
# readyReadFile(READY_FILE)
# readyReadFile(READY_FILE).to_excel("output.xlsx")
# serviceNoteReadFile(SN_FILE).to_excel("output1.xlsx")
# serviceNoteReadFile(SN_FILE)