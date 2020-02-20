
"""
Генерация итогового файла

"""
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
# import datetime
from settings import TODAY

def readAppendData(a_file):
    try:
        df = pd.read_excel(
            a_file, 
            sheet_name="Sheet1",
            header=0,
            index_col=0
        )
    except Exception as ind:
        print("serviceNoteReadFile error with file: %s - %s" % (a_file, ind))
    else:
        return df

def colorsWrite(df):
    #### Индексы
    def convert(x) -> (str, None):
        try:
            return x.strftime('%d.%m.%Y') if not pd.isnull(x) else None
        except AttributeError:
            return None
            # print(x)

    # apply the function

    df.pickup_sn_date = df.pickup_sn_date.apply(lambda x: convert(x))
    df.pickup_plan_date_f = df.pickup_plan_date_f.apply(lambda x: convert(x))
    df.pickup_date = df.pickup_date.apply(lambda x: convert(x))

    df.shipping_sn_date = df.shipping_sn_date.apply(lambda x: convert(x))
    df.shipping_plan_date_f = df.shipping_plan_date_f.apply(lambda x: convert(x))
    df.shipping_date = df.shipping_date.apply(lambda x: convert(x))

    df.design_sn_date = df.design_sn_date.apply(lambda x: convert(x))
    df.design_plan_date_f = df.design_plan_date_f.apply(lambda x: convert(x))
    df.design_date = df.design_date.apply(lambda x: convert(x))

    writer = pd.ExcelWriter('testfiles\\Documentation_Deficit_Colors.xlsx', 
        engine='xlsxwriter',
        # datetime_format='dd.mm.yyyy',
        # date_format='dd.mm.yyyy'
        )

    df.to_excel(writer, sheet_name='Документы')
    workbook  = writer.book
    worksheet = writer.sheets['Документы']

    header_format = workbook.add_format({'font_size': 10,'valign':'vcenter','bold':True,'align':'center','text_wrap':True,'color': 'white','bg_color':'#f05623','bottom':1, 'bottom_color':'#f05623'
        })
    header_col = workbook.add_format({'font_size': 10})
    
    col_KV_dates = workbook.add_format({'font_size': 10,'num_format': 'dd.mm.yyyy','color': '#215967','bg_color':'#daeef3'})
    col_KV_days = workbook.add_format({'font_size': 10,'color': '#215967','bg_color':'#daeef3','align':'right'})
    col_KV_comment = workbook.add_format({'font_size': 10,'color': '#215967','bg_color':'#daeef3','align':'left'})

    col_OS_dates = workbook.add_format({'font_size': 10,'num_format': 'dd.mm.yyyy','color': '#5b3151','bg_color':'#e4dfec'})
    col_OS_days = workbook.add_format({'font_size': 10,'color': '#5b3151','bg_color':'#e4dfec','align':'right'})
    col_OS_comment = workbook.add_format({'font_size': 10,'color': '#5b3151','bg_color':'#e4dfec','align':'left'})

    col_KD_dates = workbook.add_format({'font_size': 10,'num_format': 'dd.mm.yyyy','color': '#375623','bg_color':'#ecefe9'})
    col_KD_days = workbook.add_format({'font_size': 10,'color': '#375623','bg_color':'#ecefe9','align':'right'})
    col_KD_comment = workbook.add_format({'font_size': 10,'color': '#375623','bg_color':'#ecefe9','align':'left'})

    # header_col_dates = workbook.add_format({'font_size': 10, 'num_format': 'dd.mm.yyyy', 'color': '#215967', 'bg_color':'#daeef3'})

    first_row = [
        '№№ СЗ',
        '№ Заказа',
        'Контрагент',
        'Продукция',
        'КВ Дата начала отсчета',
        'КВ Дата выдачи по плану',
        'КВ Дата выдачи по факту',
        'КВ Разница дней',
        'КВ Комментарий',
        'ОС Дата начала отсчета',
        'ОС Дата выдачи по плану',
        'ОС Дата выдачи по факту',
        'ОС Разница дней',
        'ОС Комментарий',
        'КД Дата начала отсчета',
        'КД Дата выдачи по плану',
        'КД Дата выдачи по факту',
        'КД Разница дней',
        'КД Комментарий',
    ]

    for col_num, value in enumerate(first_row):
        try:
            worksheet.write(0, col_num + 1, value, header_format)
        except:
            pass

    rows_count = len(df.count(axis='columns'))

    worksheet.set_row(0, 50, None)

    worksheet.set_column(0,0, 0) # Первые два столбца индекс и ин_айди

    worksheet.set_column(1,1, 4, header_col) # '№№ СЗ',
    worksheet.set_column(2,2, 10, header_col) # '№ Заказа',
    worksheet.set_column(3,3, 30, header_col) # 'Контрагент',
    worksheet.set_column(4,4, 70, header_col) # 'Продукция',
    
    worksheet.set_column(5,7, 10, col_KV_dates, {'level': 2, 'hidden': True}) # 'кв даты',
    worksheet.set_column(8,8, 7, col_KV_days, {'level': 1, 'hidden': True}) # 'кв дни',
    worksheet.set_column(9,9, 15, col_KV_comment) # 'кв комментарий',

    worksheet.set_column(10,12, 10, col_OS_dates, {'level': 2, 'hidden': True}) # 'ос даты',
    worksheet.set_column(13,13, 7, col_OS_days, {'level': 1, 'hidden': True}) # 'ос дни',
    worksheet.set_column(14,14, 15, col_OS_comment) # 'ос комментарий',

    worksheet.set_column(15,17, 10, col_KD_dates, {'level': 2, 'hidden': True}) # 'кд даты',
    worksheet.set_column(18,18, 7, col_KD_days, {'level': 1, 'hidden': True}) # 'кд дни',
    worksheet.set_column(19,19, 15, col_KD_comment) # 'кд комментарий',

    # Подчеркивание внизу строки, разделяем заказы
    bottom_line = workbook.add_format({'bottom':1, 'bottom_color':'#f05623'})
    select = 'B1:T%s' % rows_count
    worksheet.conditional_format(select, {'type':'formula', 'criteria': '=$B1<>$B2', 'format': bottom_line})

    # Выделяем отрицательные для КВ
    bad = workbook.add_format({'bg_color':'#ff8596'})
    good = workbook.add_format({'bg_color':'#ebf1de'})
    alert = workbook.add_format({'bg_color':'#fcd5b4'})
    select = 'F1:J%s' % rows_count
    worksheet.conditional_format(select, {'type':'formula', 'criteria': '=$I1<0', 'format': bad})

    # worksheet.conditional_format(select, {'type':'formula', 'criteria': '=$I1>0', 'format': good})
    # worksheet.conditional_format(select, {'type':'formula', 'criteria': '=AND(ISNUMBER($I1); $I1>=0)', 'format': good})

    worksheet.freeze_panes(1, 5)

    worksheet.set_default_row(hide_unused_rows=True)
    writer.save()

if __name__ == "__main__":
    import time
    def timing():
        # Счетчик времени, таймер
        start_time = time.time()
        return lambda x: print("[{:>7.2f}с.] {}".format(time.time() - start_time, x))
    t = timing()
    df = readAppendData('testfiles\\Documentation_Deficit.xlsx')
    print(df.head(5))
    colorsWrite(df)
    t("{:>5} Конец выполнения".format(''))