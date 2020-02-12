
"""
Основной исполняемый модуль

"""
from readready import readyReadFile
from readservicenote import serviceNoteReadFile
from readindocument import worker
from readproductionplan import productionPlanReadFile
from readshedule import sheduleWorker
from settings import READY_FILE, SN_FILE, IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER, PRODUCTION_PLAN_FILE, SHEDULE_FOLDER
from writeplan import writeWorker
from appenddata import appendDataWorker
import time

if __name__ == "__main__":
    def timing():
        # Счетчик времени, таймер
        start_time = time.time()
        return lambda x: print("[{:>7.2f}с.] {}".format(time.time() - start_time, x))

    t = timing()

    readyDf = readyReadFile(READY_FILE)
    readyDf.to_excel('testfiles\\Ready.xlsx')
    t("{:>5} Готовые".format(len(readyDf)))

    serviceNoteDf = serviceNoteReadFile(SN_FILE)
    serviceNoteDf.to_excel('testfiles\\Service_Notes.xlsx')
    t("{:>5} Служебные записки".format(len(serviceNoteDf)))

    inDocumentDf = worker(IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER)
    inDocumentDf.to_excel('testfiles\\In_Documents.xlsx')
    t("{:>5} Документация".format(len(inDocumentDf)))

    productionPlanDf = productionPlanReadFile(PRODUCTION_PLAN_FILE)
    productionPlanDf.to_excel('testfiles\\Production_Plan.xlsx')
    t("{:>5} План производства".format(len(productionPlanDf)))

    correct_files, error_files = sheduleWorker(SHEDULE_FOLDER)
    correct_files.to_excel('testfiles\\Correct_Shedule_Files.xlsx')
    error_files.to_excel('testfiles\\Failure_Shedule_Files.xlsx')
    t("{:>5} Обработка графиков".format(''))

    df = writeWorker([readyDf, serviceNoteDf, productionPlanDf, inDocumentDf])
    df.to_excel('testfiles\\Plan.xlsx')
    t("{:>5} Создал план".format(''))

    df = appendDataWorker(df)
    df.to_excel('testfiles\\AppendData.xlsx')
    t("{:>5} Конец выполнения".format(''))