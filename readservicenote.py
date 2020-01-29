

"""
Служебные записки

"""
import pandas as pd

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
        return df

if __name__ == "__main__":
    from settings import SN_FILE
    print(serviceNoteReadFile(SN_FILE))