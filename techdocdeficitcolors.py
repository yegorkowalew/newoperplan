
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
    
    writer = pd.ExcelWriter('testfiles\\Documentation_Deficit_Colors.xlsx', 
        engine='xlsxwriter',
        datetime_format='dd.mm.yyyy',
        date_format='dd.mm.yyyy'
        )
    df.to_excel(writer, sheet_name='План производства')
    workbook  = writer.book
    worksheet = writer.sheets['План производства']

    header_format = workbook.add_format({'font_size': 10,'valign':'top'})
    header_col = workbook.add_format({'font_size': 10})
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

    worksheet.set_column(0,0, 0) # Первые два столбца индекс и ин_айди

    worksheet.set_column(1,1, 4, header_col) # '№№ СЗ',
    worksheet.set_column(2,2, 10, header_col) # '№ Заказа',
    worksheet.set_column(3,3, 30, header_col) # 'Контрагент',
    worksheet.set_column(4,4, 70, header_col) # 'Продукция',
    worksheet.set_column(5,7, 10, header_col) # 'Продукция',
    # worksheet.set_column(3,3, 2, header_col) # Столбец цеха
    # worksheet.set_column(4,15,0, header_col) # Столбцы дат

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