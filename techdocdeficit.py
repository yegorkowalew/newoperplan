# import itertools
import pandas as pd
# import datetime
# from datetime import datetime
from settings import TODAY

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

        'design_plan_date_f',# КД Дата выдачи по плану
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
    def get_counterparty(sn_date, plan_date_f, date, days, comment, issue):
        # 'sn_date' # Начало отсчета
        # 'plan_date_f',# КВ Дата выдачи по плану
        # 'date', # КВ Дата выдачи по факту
        # 'days', # Дней
        # 'comment' # Комментарий
        # 'issue' # Нужна документация или нет
        if issue == False:
            comment = 'Не нужны'
            date = plan_date_f
        
        days = (plan_date_f - date).days

        if days > 0:
            comment = 'Раньше на %sдн.' % days

        if days < 0:
            comment = 'Позже на %sдн.' % abs(days)

        if days == 0:
            comment = 'В день по плану'

        if pd.isnull(date):
            days = (plan_date_f - TODAY).days
            if days > 0:
                comment = 'До выдачи %sдн.' % days
                days = '! %s' % (plan_date_f - TODAY).days
            if days < 0:
                comment = 'Просрочка %sдн.' % abs(days)
                days = '! %s' % abs((plan_date_f - TODAY).days)
            if days == 0:
                comment = 'Выдача сегодня'
                days = '! %s' % (plan_date_f - TODAY).days

        return sn_date, plan_date_f, date, days, comment


    for index, row in df.iterrows():
        df.loc[index, 'pickup_sn_date'], df.loc[index, 'pickup_plan_date_f'], df.loc[index, 'pickup_date'], df.loc[index, 'pickup_days'], df.loc[index, 'pickup_comment'] = get_counterparty(row['pickup_sn_date'],row['pickup_plan_date_f'],row['pickup_date'],row['pickup_days'],row['pickup_comment'],row['pickup_issue'])
        df.loc[index, 'shipping_sn_date'], df.loc[index, 'shipping_plan_date_f'], df.loc[index, 'shipping_date'], df.loc[index, 'shipping_days'], df.loc[index, 'shipping_comment'] = get_counterparty(row['shipping_sn_date'],row['shipping_plan_date_f'],row['shipping_date'],row['shipping_days'],row['shipping_comment'],row['shipping_issue'])

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
        'design_plan_date_f', # КД Дата выдачи по плану
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
    # print(df['pickup_date'].head(20))
    df.to_excel('testfiles\\DeficitTechnicalDocumentation.xlsx')
    t("{:>5} Конец выполнения".format(''))
