
"""
Основной исполняемый модуль

"""
from readready import readyReadFile
from readservicenote import serviceNoteReadFile
from readindocument import worker
from settings import READY_FILE, SN_FILE, IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER

if __name__ == "__main__":
    readyDf = readyReadFile(READY_FILE)
    readyDf.to_excel("testfiles\\Ready.xlsx")

    serviceNoteDf = serviceNoteReadFile(SN_FILE)
    serviceNoteDf.to_excel("testfiles\\Service_Notes.xlsx")

    inDocumentDf = worker(IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER)
    inDocumentDf.to_excel('testfiles\\In_Documents.xlsx')