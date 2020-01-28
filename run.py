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
            # converters= {'Дата СЗ': pd.to_datetime},
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
        # df['sn_date'] = df['sn_date'].to_datetime(errors='ignore')
        df['sn_date'] = pd.to_datetime(df['sn_date'], errors='coerce')

        def setup_pickup_plan_date(row):
            if not pd.isnull(row['pickup_plan_date']):
                return row['pickup_plan_date']

            if not pd.isnull(row['pickup_plan_days']):
                if not pd.isnull(row['sn_date']):
                    return row['sn_date'] + pd.DateOffset(row['pickup_plan_days'])

            if not pd.isnull(row['sn_date']):
                return row['sn_date']

        def setup_shipping_plan_date(row):
            if not pd.isnull(row['shipping_plan_date']):
                return row['shipping_plan_date']

            if not pd.isnull(row['shipping_plan_days']):
                if not pd.isnull(row['sn_date']):
                    return row['sn_date'] + pd.DateOffset(row['shipping_plan_days'])

            if not pd.isnull(row['sn_date']):
                return row['sn_date']

        df['pickup_plan_date_f'] = df.apply (lambda row: setup_pickup_plan_date(row), axis=1)
        df['shipping_plan_date_f'] = df.apply (lambda row: setup_shipping_plan_date(row), axis=1)
        print(df.dtypes)
        # print(df)
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
        level_map = {'-': False}
        df['produced'] = df['status'].map(level_map)
        level_map = {'+': True}
        df['ready_status'] = df['ready'].map(level_map)
        df = df.drop(['status', 'ready'], axis=1)
        df['ready_status'] = df['ready_status'].fillna(False)
        df['produced'] = df['produced'].fillna(True)
        with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
            print(df)
        return df


# zz = serviceNoteReadFile(SN_FILE)
# print(readyReadFile(READY_FILE))
# readyReadFile(READY_FILE)
# readyReadFile(READY_FILE).to_excel("output.xlsx")
serviceNoteReadFile(SN_FILE).to_excel("output1.xlsx")
# serviceNoteReadFile(SN_FILE)