import pandas as pd

"""
Служебные записки

"""
SN_FILE = 'C:\\work\\newoperplan\\Служебные записки.xlsx'

def serviceNoteReadFile(sn_file):
    try:
        df = pd.read_excel(
            sn_file, 
            sheet_name="СЗ",
            parse_dates = [
                'Отгрузка "от"',
                'Отгрузка "до"',
                'Дата СЗ',
                'Дата СЗ Факт',
                'Комплектовочные План Дней от СЗ',
                'Комплектовочные План Вручную',
                'Отгрузочные План Дней от СЗ',
                'Отгрузочные План Вручную',
                'Конструкторская документация План Дней от СЗ',
                'Конструкторская документация План Вручную',
                'Материалы План',
            ],
            dtype={
                'ID':int,
                'Продукция':str,
                'Контрагент':str,
                '№ Заказа':str,
                '№ СЗ с изменениями':str,
            },
            index_col=0
        )

    except Exception as ind:
        # logger.error("serviceNoteReadFile error with file: %s - %s" % (sn_file, ind))
        print("serviceNoteReadFile error with file: %s - %s" % (sn_file, ind))
    else:
        df = df.rename(columns={
            'ID':'in_id',
            # 'Порядковый номер':'in_sn_no',
            'Отгрузка "от"':'shipment_from',
            'Отгрузка "до"':'shipment_before',
            'Продукция':'product_name',
            'Контрагент':'counterparty',
            '№ Заказа':'order_no',
            'Кол-во':'amount',
            '№ СЗ':'sn_no',
            '№ СЗ с изменениями':'sn_no_amended',
            'Дата СЗ':'sn_date',
            'Дата СЗ Факт':'sn_date_fact',
            'Комплектовочные План Дней от СЗ':'pickup_plan_days',
            'Комплектовочные План Вручную':'pickup_plan_date',
            'Отгрузочные План Дней от СЗ':'shipping_plan_days',
            'Отгрузочные План Вручную':'shipping_plan_date',
            'Конструкторская документация План Дней от СЗ':'design_plan_days',
            'Конструкторская документация План Вручную':'design_plan_date',
            'Материалы План':'material_plan_days',
            }
        )
        # print(df)
        # df = df.astype(object)
        # df = df.where(df.notnull(), None)
        # df_records = df.to_dict('records')
        # return df_records
        return df

"""
Готовые заказы

"""
READY_FILE = 'C:\\work\\newoperplan\\Готовые заказы.xlsx'

def readyReadFile(ready_file):
    try:
        df = pd.read_excel(
            ready_file, 
            sheet_name="Готовые заказы",
            dtype={
                'ID':int,
                'Готово':str,
                'Статус в Плане производства':str,
            },
            header=1,
            usecols = ['ID', 'Готово', 'Статус в Плане производства'],
            index_col=0
        )

    except Exception as ind:
        # logger.error("serviceNoteReadFile error with file: %s - %s" % (sn_file, ind))
        print("serviceNoteReadFile error with file: %s - %s" % (ready_file, ind))
    else:
        df = df.rename(columns={
            'ID':'inid',
            'Готово':'ready',
            'Статус в Плане производства':'status'
            }
        )
        # print(df)
        # df = df.astype(object)
        # df = df.where(df.notnull(), None)
        # df_records = df.to_dict('records')
        # return df_records
        # df['Новый'] = df['ID']
        level_map = {'-': False}
        df['not_produced'] = df['status'].map(level_map)
        level_map = {'+': True}
        df['ready_status'] = df['ready'].map(level_map)
        df = 
        # df['newColumn'] = df['status']
        return df


# zz = serviceNoteReadFile(SN_FILE)
# print(readyReadFile(READY_FILE))
readyReadFile(READY_FILE).to_excel("output.xlsx")
serviceNoteReadFile(SN_FILE).to_excel("output1.xlsx")