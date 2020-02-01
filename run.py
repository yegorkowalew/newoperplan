
"""
Основной исполняемый модуль

"""
from readready import readyReadFile
from readservicenote import serviceNoteReadFile
from readindocument import worker
from readproductionplan import productionPlanReadFile
from readshedule import sheduleWorker
from settings import READY_FILE, SN_FILE, IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER, PRODUCTION_PLAN_FILE, SHEDULE_FOLDER

if __name__ == "__main__":
    readyDf = readyReadFile(READY_FILE)
    readyDf.to_excel('testfiles\\Ready.xlsx')

    serviceNoteDf = serviceNoteReadFile(SN_FILE)
    serviceNoteDf.to_excel('testfiles\\Service_Notes.xlsx')

    inDocumentDf = worker(IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER)
    inDocumentDf.to_excel('testfiles\\In_Documents.xlsx')

    productionPlanDf = productionPlanReadFile(PRODUCTION_PLAN_FILE)
    productionPlanDf.to_excel('testfiles\\Production_Plan.xlsx')

    correct_files, error_files = sheduleWorker(SHEDULE_FOLDER)
    correct_files.to_excel('testfiles\\Correct_Shedule_Files.xlsx')
    error_files.to_excel('testfiles\\Failure_Shedule_Files.xlsx')