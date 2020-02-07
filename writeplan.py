import itertools
import pandas as pd
import datetime
from datetime import datetime

# КВ –комплектовочная ведомость, 
# ВП – ведомость покупных, 
# ОС – отгрузочная спецификация, 
# КД – конструкторская документация

shop_symbol = ['Ц', 'М', 'К', 'З', 'КВ', 'ОС']
n_shop_symbol = itertools.cycle(shop_symbol)

def findMinMaxDate(df):
    maxDates = []
    minDates = []
    for column in list(df.columns.values):
        try:
            maxDates.append(df[column].max())
        except:
            pass
        try:
            minDates.append(df[column].min())
        except:
            pass
    max_date = pd.to_datetime(pd.Series(maxDates), errors='coerce').max()
    min_date = pd.to_datetime(pd.Series(minDates), errors='coerce').min()
    return min_date, max_date

def create_dataframe(dflist):
    df = pd.concat(dflist, axis=1, sort=False)
    df = df.loc[(df['ready_status'] == False) & (df['produced'] == True)]

    # Нужно создать три столбца:
    # days_compl_plan "Дней на выполнение по плану. Дата выполнения по плану минус дата служебной записки."
    # days_compl_go "Дней на выполнение прошло. Устанавливается в том случае, если еще не выполнено. Сегодняшняя дата минус дата начала выполнения."
    # days_compl_execution "Дней на выполнение потрачено. Устанавливается в том случае, если задача выполнена. Дата начала выполнения - дата окончания выполнения"

    # df['days_compl_pickup_plan'] = 'www'
    # days_compl_pickup_plan "Дней на выполнение комплектовочных по плану.
    return df

def writeWorker(dflist):
    df = create_dataframe(dflist)
    min_date, max_date = findMinMaxDate(df)

    # df.to_excel('testfiles\\out.xlsx')

    def get_product_name(row):
        if not pd.isnull(row['product_name']):
            return row['product_name']
        else:
            return 'Не установлено'

    def get_info(row):
        if not pd.isnull(row['order_no']):
            order_no = str(row['order_no'])
        else:
            order_no = 'Не установлено'

        if not pd.isnull(row['sn_no']):
            sn_no = str(int(row['sn_no']))
        else:
            sn_no = 'Не установлено'

        if not pd.isnull(row['amount']):
            amount = str(row['amount'])
        else:
            amount = 'Не установлено'
        return '№Зк: %s, №СЗ: %s, Кол-во: %s' % (order_no, sn_no, amount)

    def get_counterparty(row):
        if not pd.isnull(row['counterparty']):
            return 'Заказчик: %s' % (row['counterparty'])
        else:
            return 'Заказчик: Не указан'

    def get_shipment_info(row):
        if not pd.isnull(row['shipment_from']) and not pd.isnull(row['shipment_before']):
            shipment_from_days = (row['shipment_from'] - pd.Timestamp.today()).days
            shipment_before_days = (row['shipment_before'] - pd.Timestamp.today()).days
            if shipment_before_days > 0:
                shipment = '%s-%sдн.' % (shipment_from_days, shipment_before_days)
            else:
                shipment = 'Просрочка: %sдн.' % shipment_before_days
            return 'Отгрузка: %s-%s (%s)' % (
                row['shipment_from'].strftime("%d.%m.%Y"), 
                row['shipment_before'].strftime("%d.%m.%Y"),
                shipment
                )
        if pd.isnull(row['shipment_from']) and not pd.isnull(row['shipment_before']):
            shipment_before_days = (row['shipment_before'] - pd.Timestamp.today()).days
            if shipment_before_days > 0:
                shipment = '%sдн.' % (shipment_before_days)
            else:
                shipment = 'Просрочка: %sдн.' % shipment_before_days
            return 'Отгрузка: %s (%s)' % (
                row['shipment_before'].strftime("%d.%m.%Y"),
                shipment
                )
        if pd.isnull(row['shipment_from']) and pd.isnull(row['shipment_before']):
            return 'Отгрузка: Не установлено'

    def get_pickup_doc_info(row):
        returnstr = 'КВ: %s' % row['pickup_plan_date_f'].strftime("%d.%m.%Y") 
        if row['pickup_issue'] == False:
            returnstr = '%s (Не нужны, %s)' % (returnstr, row['dispatcher_pickup_issue'])
            return returnstr
        else:
            if not pd.isnull(row['pickup_date']):
                # Дата прихода документации установлена
                pickup_days = (row['pickup_plan_date_f'] - row['pickup_date']).days
                pickup_str = '%sдн.' % pickup_days
                if pickup_days > 0:
                    return '%s (Получено на %s раньше)' % (returnstr, pickup_str)
                else:
                    return '%s (Получено на %s позже)' % (returnstr, pickup_str)
            else:
                pickup_days = (row['pickup_plan_date_f'] - pd.Timestamp.today()).days
                if pickup_days > 0:
                    return '%s (Осталось %sдн.)' % (returnstr, pickup_days)
                else:
                    return '%s (Просрочено %sдн.)' % (returnstr, pickup_days)

    def get_shipping_doc_info(row):
        # return 'ОС: %s' % row['shipping_plan_date_f'].strftime("%d.%m.%Y")
        returnstr = 'ОС: %s' % row['shipping_plan_date_f'].strftime("%d.%m.%Y") 
        if row['shipping_issue'] == False:
            returnstr = '%s (Не нужны, %s)' % (returnstr, row['dispatcher_shipping_issue'])
            return returnstr
        else:
            if not pd.isnull(row['shipping_date']):
                # Дата прихода документации установлена
                shipping_days = (row['shipping_plan_date_f'] - row['shipping_date']).days
                shipping_str = '%sдн.' % shipping_days
                if shipping_days > 0:
                    return '%s (Получено на %s раньше)' % (returnstr, shipping_str)
                else:
                    return '%s (Получено на %s позже)' % (returnstr, shipping_str)
            else:
                shipping_days = (row['shipping_plan_date_f'] - pd.Timestamp.today()).days
                if shipping_days > 0:
                    return '%s (Осталось %sдн.)' % (returnstr, shipping_days)
                else:
                    return '%s (Просрочено %sдн.)' % (returnstr, shipping_days)

    def get_work_start(row, shop_letter):
        if shop_letter == 'К':
            return row['101_start_date']
        if shop_letter == 'М':
            return row['102_start_date']
        if shop_letter == 'Ц':
            return row['104_start_date']
        if shop_letter == 'З':
            return row['107_start_date']
        if shop_letter == 'КВ':
            return row['sn_date']
        if shop_letter == 'ОС':
            return row['sn_date']
        
    def get_work_end(row, shop_letter):
        if shop_letter == 'К':
            return row['101_end_date']
        if shop_letter == 'М':
            return row['102_end_date']
        if shop_letter == 'Ц':
            return row['104_end_date']
        if shop_letter == 'З':
            return row['107_end_date']
        if shop_letter == 'КВ':
            return row['pickup_plan_date_f']
        if shop_letter == 'ОС':
            return row['shipping_plan_date_f']

    def get_work_zinc(row, shop_letter):
        if shop_letter == 'К':
            return row['101_zinc']
        if shop_letter == 'М':
            return row['102_zinc']
        if shop_letter == 'Ц':
            return row['104_zinc']
        if shop_letter == 'З':
            return row['107_zinc']

    def get_work_rubberizing(row):
        return None

    def get_order_finish(row):
        if row['produced'] and row['ready_status']:
            return row['ready_date']
        # if row['produced'] and not row['ready_status']:
        #     return 'Изготавливается и не готово'

    col_1 = [] # Строка с названием товара
    col_2 = [] # Cтрока с буквой цеха
    col_3 = [] # Идентификатор заказа
    col_4 = [] # Дата начала работ
    col_5 = [] # Дата планового окончания работ
    col_6 = [] # Дата планового цинкования
    col_7 = [] # Дата планового обрезинивания
    col_8 = [] # Дата планового старта заказа (Дата служебной записки)
    col_9 = [] # Дата планового плановой отгрузки ОТ
    col_10 = [] # Дата планового плановой отгрузки ДО
    col_11 = [] # Дата планового поступления материала
    col_12 = [] # Дата завершения заказа по факту
    for index, row in df.iterrows():
        col_1.append(get_product_name(row))
        shop = next(n_shop_symbol)
        col_2.append(shop)
        col_3.append(index)
        col_4.append(get_work_start(row, shop))
        col_5.append(get_work_end(row, shop))
        col_6.append(get_work_zinc(row, shop))
        col_7.append(get_work_rubberizing(row))
        col_8.append(row['sn_date'])
        col_9.append(row['shipment_from'])
        col_10.append(row['shipment_before'])
        col_11.append(row['material_plan_date'])
        col_12.append(get_order_finish(row))

        col_1.append(get_info(row))
        shop = next(n_shop_symbol)
        col_2.append(shop)
        col_3.append(index)
        col_4.append(get_work_start(row, shop))
        col_5.append(get_work_end(row, shop))
        col_6.append(get_work_zinc(row, shop))
        col_7.append(get_work_rubberizing(row))
        col_8.append(row['sn_date'])
        col_9.append(row['shipment_from'])
        col_10.append(row['shipment_before'])
        col_11.append(row['material_plan_date'])
        col_12.append(get_order_finish(row))

        col_1.append(get_counterparty(row))
        shop = next(n_shop_symbol)
        col_2.append(shop)
        col_3.append(index)
        col_4.append(get_work_start(row, shop))
        col_5.append(get_work_end(row, shop))
        col_6.append(get_work_zinc(row, shop))
        col_7.append(get_work_rubberizing(row))
        col_8.append(row['sn_date'])
        col_9.append(row['shipment_from'])
        col_10.append(row['shipment_before'])
        col_11.append(row['material_plan_date'])
        col_12.append(get_order_finish(row))

        col_1.append(get_shipment_info(row))
        shop = next(n_shop_symbol)
        col_2.append(shop)
        col_3.append(index)
        col_4.append(get_work_start(row, shop))
        col_5.append(get_work_end(row, shop))
        col_6.append(get_work_zinc(row, shop))
        col_7.append(get_work_rubberizing(row))
        col_8.append(row['sn_date'])
        col_9.append(row['shipment_from'])
        col_10.append(row['shipment_before'])
        col_11.append(row['material_plan_date'])
        col_12.append(get_order_finish(row))

        col_1.append(get_pickup_doc_info(row))
        shop = next(n_shop_symbol)
        col_2.append(shop)
        col_3.append(index)
        col_4.append(get_work_start(row, shop))
        col_5.append(get_work_end(row, shop))
        col_6.append(get_work_zinc(row, shop))
        col_7.append(get_work_rubberizing(row))
        col_8.append(row['sn_date'])
        col_9.append(row['shipment_from'])
        col_10.append(row['shipment_before'])
        col_11.append(row['material_plan_date'])
        col_12.append(get_order_finish(row))

        col_1.append(get_shipping_doc_info(row))
        shop = next(n_shop_symbol)
        col_2.append(shop)
        col_3.append(index)
        col_4.append(get_work_start(row, shop))
        col_5.append(get_work_end(row, shop))
        col_6.append(get_work_zinc(row, shop))
        col_7.append(get_work_rubberizing(row))
        col_8.append(row['sn_date'])
        col_9.append(row['shipment_from'])
        col_10.append(row['shipment_before'])
        col_11.append(row['material_plan_date'])
        col_12.append(get_order_finish(row))

    frame = {
        'in_id': col_3, 
        'product': col_1, 
        'shop': col_2, 
        'order_plan_start':col_8,
        'work_start':col_4, 
        'work_end':col_5, 
        'zinc':col_6,
        'rubberizing':col_7,
        'material':col_11,
        'order_plan_shipment_from':col_9,
        'order_plan_shipment_before':col_10,
        'order_finish':col_12,
        }
        
    result = pd.DataFrame(frame) 
    result.to_excel('testfiles\\Plan.xlsx')

if __name__ == "__main__":
    from settings import READY_FILE, SN_FILE, IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER, PRODUCTION_PLAN_FILE, SHEDULE_FOLDER
    from readready import readyReadFile
    from readservicenote import serviceNoteReadFile
    from readproductionplan import productionPlanReadFile
    from readindocument import worker

    import time
    def timing():
        # Счетчик времени, таймер
        start_time = time.time()
        return lambda x: print("[{:>7.2f}с.] {}".format(time.time() - start_time, x))

    t = timing()

    readyDf = readyReadFile(READY_FILE)
    t("{:>5} Готовые".format(len(readyDf)))

    serviceNoteDf = serviceNoteReadFile(SN_FILE)
    t("{:>5} Служебные записки".format(len(serviceNoteDf)))

    productionPlanDf = productionPlanReadFile(PRODUCTION_PLAN_FILE)
    t("{:>5} План производства".format(len(productionPlanDf)))

    inDocumentDf = worker(IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER)
    t("{:>5} Документация".format(len(inDocumentDf)))

    writeWorker([readyDf, serviceNoteDf, productionPlanDf, inDocumentDf])
    t("{:>5} Создал план".format(''))

    t("{:>5} Конец выполнения".format(''))