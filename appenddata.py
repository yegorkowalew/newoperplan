import itertools
import pandas as pd
import datetime
from datetime import datetime

def setup_order_plan_start(df, row_index, dates_base):
    order_plan_start = df.loc[row_index, 'order_plan_start']
    if not pd.isnull(order_plan_start):
        order_plan_start = pd.to_datetime(order_plan_start, format='%Y-%m-%d %H:%M:%S.%f')
        dates_base = dates_base[dates_base['dates'] == order_plan_start]
        df.loc[row_index, dates_base.index[0]] = 'СЗ'

def setup_material(df, row_index, dates_base):
    order_plan_start = df.loc[row_index, 'material']
    if not pd.isnull(order_plan_start):
        order_plan_start = pd.to_datetime(order_plan_start, format='%Y-%m-%d %H:%M:%S.%f')
        dates_base = dates_base[dates_base['dates'] == order_plan_start]
        df.loc[row_index, dates_base.index[0]] = 'm'

def setup_work_plan(df, row_index, dates_base):
    work_start = df.loc[row_index, 'work_start']
    work_end_plan = df.loc[row_index, 'work_end_plan']
    work_end_fact = df.loc[row_index, 'work_end_fact']
    today = pd.to_datetime('2020-02-10 00:00:00', format='%Y-%m-%d %H:%M:%S.%f')
    shop = df.loc[row_index, 'shop']
    if work_start:
        work_start = pd.to_datetime(work_start, format='%Y-%m-%d %H:%M:%S.%f')
        work_end_plan = pd.to_datetime(work_end_plan, format='%Y-%m-%d %H:%M:%S.%f')
        dates_base_list = dates_base[(dates_base['dates'] >= work_start) & (dates_base['dates'] <= work_end_plan)]
        df.loc[row_index, dates_base_list.index.to_list()] = shop
        if (work_end_plan - work_start).days > 1:
            df.loc[row_index, dates_base_list.index.to_list()[-1]+1] = (work_end_plan - work_start).days+1
    if not pd.isnull(work_end_fact):
        work_end_fact = pd.to_datetime(work_end_fact, format='%Y-%m-%d %H:%M:%S.%f')
        dates_base_list = dates_base[(dates_base['dates'] > work_end_plan) & (dates_base['dates'] <= work_end_fact)]
        # print(dates_base_list.index, '<<<<<')
        df.loc[row_index, dates_base_list.index[1:]] = 'X'
        if (work_end_fact - work_end_plan).days > 1:
            df.loc[row_index, dates_base_list.index[-1]+1] = (work_end_fact - work_end_plan).days
    if pd.isnull(work_end_fact):
        # work_end_fact = pd.to_datetime(work_end_fact, format='%Y-%m-%d %H:%M:%S.%f')
        work_end_plan = pd.to_datetime(work_end_plan, format='%Y-%m-%d %H:%M:%S.%f')
        # today
        if (today - work_end_plan).days > 1:
            # print(today, ' - ' ,work_end_plan, ' - ', (today - work_end_plan).days)
            dates_base_list = dates_base[(dates_base['dates'] > work_end_plan) & (dates_base['dates'] <= today)]
            # print(dates_base_list.index.to_list())
            df.loc[row_index, dates_base_list.index[1:]] = '>'
            df.loc[row_index, dates_base_list.index[-1]+1] = (today - work_end_plan).days
            # print('----------------------')


def setup_shipment(df, row_index, dates_base):
    order_plan_shipment_from = df.loc[row_index, 'order_plan_shipment_from']
    order_plan_shipment_before = df.loc[row_index, 'order_plan_shipment_before']
    order_plan_start = df.loc[row_index, 'order_plan_start']
    if order_plan_start:
        days_count = True
    if not pd.isnull(order_plan_shipment_from):
        if days_count:
            order_plan_shipment_from_days = (order_plan_shipment_from - order_plan_start).days
            order_plan_shipment_before_days = (order_plan_shipment_before - order_plan_start).days
        order_plan_shipment_from = pd.to_datetime(order_plan_shipment_from, format='%Y-%m-%d %H:%M:%S.%f')
        order_plan_shipment_before = pd.to_datetime(order_plan_shipment_before, format='%Y-%m-%d %H:%M:%S.%f')
        dates_base_list = dates_base[(dates_base['dates'] >= order_plan_shipment_from) & (dates_base['dates'] <= order_plan_shipment_before)]
        df.loc[row_index, dates_base_list.index.to_list()] = 'S'
        if order_plan_shipment_from_days > 1:
            df.loc[row_index, dates_base_list.index[0]+1] = order_plan_shipment_from_days+1
            df.loc[row_index, dates_base_list.index[-1]+1] = order_plan_shipment_before_days+1
    elif not pd.isnull(order_plan_shipment_before):
        if days_count:
            order_plan_shipment_before_days = (order_plan_shipment_before - order_plan_start).days
        order_plan_shipment_before = pd.to_datetime(order_plan_shipment_before, format='%Y-%m-%d %H:%M:%S.%f')
        dates_base_list = dates_base[dates_base['dates'] == order_plan_shipment_before]
        df.loc[row_index, dates_base_list.index[0]] = 'S'
        df.loc[row_index, dates_base_list.index[0]+1] = order_plan_shipment_before_days+1

def appendDataWorker(df):
    dates_base = pd.DataFrame({
        'dates': pd.to_datetime(df.iloc[0], format='%Y-%m-%d %H:%M:%S.%f')
    })
    dates_base['dates'] = pd.to_datetime(dates_base['dates'])
    from itertools import islice
    
    for index, row in islice(df.iterrows(), 2, None):
        if not pd.isnull(row['order_plan_start']):
            # Если не установлена дата СЗ, то запускать установку дней не будем
            setup_order_plan_start(df, index, dates_base)
            setup_shipment(df, index, dates_base)
            setup_work_plan(df, index, dates_base)
            setup_material(df, index, dates_base)
        else:
            print('Дата начала работ (СЗ) не установлена')
    
    return df

if __name__ == "__main__":
    from writeplan import dates_to_header
    df = pd.DataFrame({
        'in_id':[1285, 1285, 1285, 1285, 1285, 1285, 1286, 1286, 1286, 1286, 1286, 1286],
        'product':['СМВУА.73.06.К45.В12', '№Зк: 2311879, №СЗ: 369, Кол-во: 1.0', 'Заказчик: ФГ "ХЛІБ-АГРО"', 'Отгрузка: 12.02.2020-19.02.2020 (3-10дн.)', 'КВ: 02.12.2019 (Получено на -2дн. позже)', 'ОС: 02.12.2019 (Получено на -9дн. позже)', 'СМВУ.55.08.К45.В12.А', '№Зк: 2311880, №СЗ: 369, Кол-во: 1.0', 'Заказчик: ФГ "ХЛІБ-АГРО"', 'Отгрузка: 12.02.2020-19.02.2020 (3-10дн.)', 'КВ: 02.12.2019 (Получено на -8дн. позже)', 'ОС: 02.12.2019 (Получено на -10дн. позже)'],
        'shop':['Ц', 'М', 'К', 'З', 'КВ', 'ОС', 'Ц', 'М', 'К', 'З', 'КВ', 'ОС'],
        'order_plan_start':['', '2019-11-11 00:00:00', '2019-11-11 00:00:00', '2019-11-11 00:00:00', '2019-11-11 00:00:00', '2019-11-11 00:00:00', '2019-11-12 00:00:00', '2019-11-12 00:00:00', '2019-11-12 00:00:00', '2019-11-12 00:00:00', '2019-11-12 00:00:00', '2019-11-12 00:00:00'],
        'work_start':['2019-12-02 00:00:00', '2019-12-02 00:00:00', '2019-12-09 00:00:00', '2019-12-02 00:00:00', '2019-11-12 00:00:00', '2019-11-12 00:00:00', '2019-12-02 00:00:00', '2019-12-02 00:00:00', '2019-12-09 00:00:00', '2019-12-02 00:00:00', '2019-11-12 00:00:00', '2019-11-12 00:00:00'],
        'work_end_plan':['2020-02-07 00:00:00', '2019-12-13 00:00:00', '2020-02-07 00:00:00', '2019-12-13 00:00:00', '2019-12-02 00:00:00', '2019-12-02 00:00:00', '2020-02-07 00:00:00', '2019-12-13 00:00:00', '2020-02-07 00:00:00', '2019-12-13 00:00:00', '2019-12-02 00:00:00', '2019-12-02 00:00:00'],
        'work_end_fact':['', '', '', '', '2019-12-04 00:00:00', '2019-12-11 00:00:00', '', '', '', '', '2019-12-10 00:00:00', '2019-12-12 00:00:00'],
        'zinc':['', '', '', '', '', '', '', '', '', '', '', ''],
        'rubberizing':['', '', '', '', '', '', '', '', '', '', '', ''],
        'material':['2019-12-22 00:00:00', '2019-12-22 00:00:00', '2019-12-22 00:00:00', '2019-12-22 00:00:00', '2019-12-22 00:00:00', '2019-12-22 00:00:00', '2019-12-27 00:00:00', '2019-12-27 00:00:00', '2019-12-27 00:00:00', '2019-12-27 00:00:00', '2019-12-27 00:00:00', '2019-12-27 00:00:00'],
        'order_plan_shipment_from':['2020-02-12 00:00:00', '', '2020-02-12 00:00:00', '2020-02-12 00:00:00', '2020-02-12 00:00:00', '2020-02-12 00:00:00', '2020-02-12 00:00:00', '2020-02-12 00:00:00', '2020-02-12 00:00:00', '2020-02-12 00:00:00', '2020-02-12 00:00:00', '2020-02-12 00:00:00'],
        'order_plan_shipment_before':['2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00', '2020-02-19 00:00:00'],
        'order_finish':['', '', '', '', '', '', '', '', '', '', '', ''],
        })
    df['order_plan_start'] =  pd.to_datetime(df['order_plan_start'], format='%Y-%m-%d %H:%M:%S.%f')
    df['work_start'] =  pd.to_datetime(df['work_start'], format='%Y-%m-%d %H:%M:%S.%f')
    df['work_end_plan'] =  pd.to_datetime(df['work_end_plan'], format='%Y-%m-%d %H:%M:%S.%f')
    df['work_end_fact'] =  pd.to_datetime(df['work_end_fact'], format='%Y-%m-%d %H:%M:%S.%f')
    df['zinc'] =  pd.to_datetime(df['zinc'], format='%Y-%m-%d %H:%M:%S.%f')
    df['rubberizing'] =  pd.to_datetime(df['rubberizing'], format='%Y-%m-%d %H:%M:%S.%f')
    df['material'] =  pd.to_datetime(df['material'], format='%Y-%m-%d %H:%M:%S.%f')
    df['order_plan_shipment_from'] =  pd.to_datetime(df['order_plan_shipment_from'], format='%Y-%m-%d %H:%M:%S.%f')
    df['order_plan_shipment_before'] =  pd.to_datetime(df['order_plan_shipment_before'], format='%Y-%m-%d %H:%M:%S.%f')
    df['order_finish'] =  pd.to_datetime(df['order_finish'], format='%Y-%m-%d %H:%M:%S.%f')
    df = dates_to_header(df)

    appendDataWorker(df)

    df.to_excel('testfiles\\AppendData.xlsx')