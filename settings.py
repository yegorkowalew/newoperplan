from datetime import datetime
import os
# TODAY = datetime.strptime('2020-02-14', "%Y-%m-%d")
TODAY = datetime.now()

## Файл служебных записок
SN_FILE = 'C:\\work\\newoperplan\\testfiles\\Служебные записки.xlsx'

## Файл с готовыми заказами
READY_FILE = 'C:\\work\\newoperplan\\testfiles\\Готовые заказы.xlsx'

## Папка в которой диспетчера отмечают входящую документацию и имя файла который нужно парсить
IN_DOCUMENT_FOLDER = 'C:\\work\\newoperplan\\testfiles\\График документации'
IN_DOCUMENT_FILE = 'График документации v1.xlsx'

## План производства
PRODUCTION_PLAN_FILE = 'C:\\work\\newoperplan\\testfiles\\План производства.xlsx' # Файл плана производства

## Папка графиков ПДО
SHEDULE_FOLDER = 'C:\\work\\newoperplan\\testfiles\\Графики ПДО'

## Папка "Учет конструкторской документации"
TECH_DOC_FOLDER = 'C:\\work\\newoperplan\\testfiles\\Учет конструкторской документации'

## Папка "Учет конструкторской документации" - База дефицитов
TECH_DOC_BASE_FILE = os.path.join(TECH_DOC_FOLDER, 'База дефицитов', 'База.xlsx')

## Папка "Учет конструкторской документации" - Дефициты
TECH_DOC_DEFICIT_FOLDER = os.path.join(TECH_DOC_FOLDER, 'Дефициты')

## Папка "Учет конструкторской документации" - Ежедневные отчеты
TECH_DOC_DAILY_REPORT_FOLDER = os.path.join(TECH_DOC_FOLDER, 'Ежедневные отчеты')

legend = {
    'Значение':'Описание',
    'Ц':'Цех. 104, ЦВС',
    'М':'Цех. 102, Механический цех',
    'К':'Цех. 101, Котельный цех',
    'З':'Цех. 107, Заготовительный цех',
    'КВ':'Документ. Комплектовочная ведомость',
    'ОС':'Документ. Отгрузочная спецификация',
    'X':'Событие. Событие состоялось, но позже даты по плану',
    '>':'Событие. Событие еще не состоялось',
    'm':'Документ. Дата планового обеспечения материалами',
    'S':'Событие. Отгрузка заказа',
}