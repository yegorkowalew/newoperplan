# import itertools
import pandas as pd
# import datetime
# from datetime import datetime

def create_dataframe(dflist):
    df = pd.concat(dflist, axis=1, sort=False)
    df = df.loc[(df['ready_status'] == False) & (df['produced'] == True)]
        # №№СЗ
        # № Заказа
        # Контрагент
        # Продукция
        # Отгрузка
        # КВ Дата начала отсчета
        # КВ Дата выдачи по плану
        # КВ Дата выдачи по факту
        # КВ Разница дней
        # КВ Комментарий
        # ОС Дата начала отсчета
        # ОС Дата выдачи по плану
        # ОС Дата выдачи по факту
        # ОС Разница дней
        # ОС Комментарий
        # КД Дата начала отсчета
        # КД Дата выдачи по плану
        # КД Дата выдачи по факту
        # КД Разница дней
        # КД Комментарий
        
    df = df[[
        'sn_no', # №№СЗ
        'order_no', # № Заказа
        'counterparty', # Контрагент
        'product_name', # Продукция
        # 'shipment_from',
        'shipment_before', # Отгрузка
        # 'amount',
        'sn_date', # Дата начала комплектовочных, отгрузочных, чертежей,
        
        # КВ Дата начала отсчета 'sn_date'
        'pickup_plan_date_f',# КВ Дата выдачи по плану
        'pickup_date', # КВ Дата выдачи по факту
        # 'dispatcher_pickup_date', # Диспетчер отметевший дату выдачи по факту, пока не нужен
        'pickup_issue', # Нужна документация или нет
        'dispatcher_pickup_issue', # Диспетчер отметивший нужна документация или нет
        # КВ Разница дней
        # КВ Комментарий

        # ОС Дата начала отсчета 'sn_date'
        'shipping_plan_date_f', # ОС Дата выдачи по плану
        'shipping_date', # ОС Дата выдачи по факту
        'dispatcher_shipping_date', # Диспетчер отметевший дату выдачи по факту, пока не нужен
        'shipping_issue', # Нужна документация или нет
        'dispatcher_shipping_issue', # Диспетчер отметивший нужна документация или нет
        # ОС Разница дней
        # ОС Комментарий
    ]].copy()
    df = df.dropna(subset=['sn_no'])
    return df

def worker_tech_doc(df):
    # Создаю столбцы с датами начала отсчета
    df['pickup_sn_date'] = df['sn_date']
    df['shipping_sn_date'] = df['sn_date']
    df['design_sn_date'] = df['sn_date']
    
    # Столбцы с разницей дней
    df['pickup_days'] = None
    df['shipping_days'] = None
    df['design_days'] = None

    # Столбцы с комментарием
    df['pickup_comment'] = None
    df['shipping_comment'] = None
    df['design_comment'] = None
    df['pickup_issue'] = df['pickup_issue'].fillna(True).astype(bool)
    # Если не нужны, ставим в факт дату сз, в коммент пишем "Не нужны"
    def get_counterparty(pickup_sn_date, pickup_plan_date_f, pickup_date, pickup_days, pickup_comment, pickup_issue):
        # 'pickup_sn_date' # Начало отсчета
        # 'pickup_plan_date_f',# КВ Дата выдачи по плану
        # 'pickup_date', # КВ Дата выдачи по факту
        # 'pickup_days', # Дней
        # 'pickup_comment' # Комментарий
        # 'pickup_issue' # Нужна документация или нет
        # if not pd.isnull(row['counterparty']):
        #     return 'Заказчик: %s' % (row['counterparty'])
        # else:
        #     return 'Заказчик: Не указан'
        if pickup_issue == False:
            pickup_comment = 'Не нужны'
            pickup_date = pickup_sn_date
            pickup_days = pickup_sn_date - pickup_sn_date 
        # pickup_days = pickup_plan_date_f - pickup_sn_date
        return pickup_sn_date, pickup_plan_date_f, pickup_date, pickup_days, pickup_comment
        # return 1, 2, 3, 4, 5


    for index, row in df.iterrows():
        # if index == 1285:
        # row['pickup_sn_date'],row['pickup_plan_date_f'],row['pickup_date'],row['pickup_days'],row['pickup_comment'] = get_counterparty(row['pickup_sn_date'],row['pickup_plan_date_f'],row['pickup_date'],row['pickup_days'],row['pickup_comment'],row['pickup_issue'])
            # print(row)
        df.loc[index, 'pickup_sn_date'], df.loc[index, 'pickup_plan_date_f'], df.loc[index, 'pickup_date'], df.loc[index, 'pickup_days'], df.loc[index, 'pickup_comment'] = get_counterparty(row['pickup_sn_date'],row['pickup_plan_date_f'],row['pickup_date'],row['pickup_days'],row['pickup_comment'],row['pickup_issue'])

    # Пересобираю df в правильном порядке столбцов
    df = df[[
        'sn_no', # №№СЗ
        'order_no', # № Заказа
        'counterparty', # Контрагент
        'product_name', # Продукция
        # 'shipment_from',
        # 'shipment_before', # Отгрузка
        # 'amount',
        # 'sn_date', # Дата начала комплектовочных, отгрузочных, чертежей,
        
        'pickup_sn_date', # КВ Дата начала отсчета 'sn_date'
        'pickup_plan_date_f',# КВ Дата выдачи по плану
        'pickup_date', # КВ Дата выдачи по факту
        'pickup_days', # Разница в днях
        'pickup_comment', # Комментарий
        # 'dispatcher_pickup_date', # Диспетчер отметевший дату выдачи по факту, пока не нужен
        # 'pickup_issue', # Нужна документация или нет
        # 'dispatcher_pickup_issue', # Диспетчер отметивший нужна документация или нет

        'shipping_sn_date', # ОС Дата начала отсчета 'sn_date'
        'shipping_plan_date_f', # ОС Дата выдачи по плану
        'shipping_date', # ОС Дата выдачи по факту
        'shipping_days', # Разница в днях
        'shipping_comment', # Комментарий
        # 'dispatcher_shipping_date', # Диспетчер отметевший дату выдачи по факту, пока не нужен
        # 'shipping_issue', # Нужна документация или нет
        # 'dispatcher_shipping_issue', # Диспетчер отметивший нужна документация или нет

        'design_sn_date',
        'design_days', # Разница в днях
        'design_comment', # Комментарий
    ]]

    return df

if __name__ == "__main__":
    from settings import READY_FILE, SN_FILE, IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER, PRODUCTION_PLAN_FILE, SHEDULE_FOLDER
    from readready import readyReadFile
    from readservicenote import serviceNoteReadFile
    from readproductionplan import productionPlanReadFile
    from readindocument import worker
    from appenddata import appendDataWorker

    import time
    def timing():
        # Счетчик времени, таймер
        start_time = time.time()
        return lambda x: print("[{:>7.2f}с.] {}".format(time.time() - start_time, x))
    
    t = timing() #[  76.51с.]       Конец выполнения
    
    from multiprocessing import Pool
    from multiprocessing.dummy import Pool as ThreadPool
    pool = ThreadPool()

    results = []
    result_objects = [
        pool.apply_async(readyReadFile, (READY_FILE,)),
        pool.apply_async(serviceNoteReadFile, (SN_FILE,)),
        pool.apply_async(productionPlanReadFile, (PRODUCTION_PLAN_FILE,)),
        pool.apply_async(worker, (IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER,)),
        ]
    
    df = create_dataframe([result_objects[0].get(), result_objects[1].get(), result_objects[2].get(), result_objects[3].get()])
    df = worker_tech_doc(df)
    df.to_excel('testfiles\\DeficitTechnicalDocumentation.xlsx')
    t("{:>5} Конец выполнения".format(''))
