import itertools
import pandas as pd
import datetime
from datetime import datetime

# КВ –комплектовочная ведомость, 
# ВП – ведомость покупных, 
# ОС –отгрузочная спецификация, 
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

    df['days_compl_pickup_plan'] = 'www'
    # days_compl_pickup_plan "Дней на выполнение комплектовочных по плану.
    return df

def writeWorker(dflist):
    df = create_dataframe(dflist)
    min_date, max_date = findMinMaxDate(df)

    df.to_excel('testfiles\\out.xlsx')

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
        return 'ОС: %s' % row['shipping_plan_date_f'].strftime("%d.%m.%Y")

    col_1 = []
    col_2 = []
    col_3 = []
    col_4 = []
    for index, row in df.iterrows():
        col_1.append(get_product_name(row))
        col_2.append(next(n_shop_symbol))
        col_3.append(index)

        col_1.append(get_info(row))
        col_2.append(next(n_shop_symbol))
        col_3.append(index)

        col_1.append(get_counterparty(row))
        col_2.append(next(n_shop_symbol))
        col_3.append(index)

        col_1.append(get_shipment_info(row))
        col_2.append(next(n_shop_symbol))
        col_3.append(index)

        col_1.append(get_pickup_doc_info(row))
        col_2.append(next(n_shop_symbol))
        col_3.append(index)

        col_1.append(get_shipping_doc_info(row))
        col_2.append(next(n_shop_symbol))
        col_3.append(index)

    frame = { 'Index': col_3, 'Product': col_1, 'Shop': col_2}
    result = pd.DataFrame(frame) 
    result.to_excel('testfiles\\Plan.xlsx')

if __name__ == "__main__":
    from settings import READY_FILE, SN_FILE, IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER, PRODUCTION_PLAN_FILE, SHEDULE_FOLDER
    from readready import readyReadFile
    from readservicenote import serviceNoteReadFile
    from readproductionplan import productionPlanReadFile
    from readindocument import worker

    readyDf = readyReadFile(READY_FILE)
    serviceNoteDf = serviceNoteReadFile(SN_FILE)
    productionPlanDf = productionPlanReadFile(PRODUCTION_PLAN_FILE)
    inDocumentDf = worker(IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER)
    writeWorker([readyDf, serviceNoteDf, productionPlanDf, inDocumentDf])