
"""
Генерация итогового файла

"""
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

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
    def get_first_date():
        header_list = list(df.columns.values)
        for itm in header_list:
            try:
                return header_list.index(int(itm))+1
            except:
                pass

    # Последняя строка
    rows_count = len(df.count(axis='columns'))
    columns_count = len(df.count(axis='rows'))
    first_date = get_first_date()
    writer = pd.ExcelWriter('testfiles\\AppendDataColors.xlsx', 
        engine='xlsxwriter',
        datetime_format='dd.mm.yyyy',
        date_format='dd.mm.yyyy'
        )
    df.to_excel(writer, sheet_name='План производства')
    workbook  = writer.book
    worksheet = writer.sheets['План производства']
    cfmt_all = workbook.add_format({'font_size': 10})
    cfmt_header_col = workbook.add_format({'font_size': 10})
    
    cfmt_SZ = workbook.add_format({'color': '#f0800a', 'bg_color':'#fde9d9'})
    cfmt_C = workbook.add_format({'color': '#d6a300', 'bg_color':'#fff2c8'})
    cfmt_M = workbook.add_format({'color': '#548235', 'bg_color':'#e2efda'})
    cfmt_K = workbook.add_format({'color': '#2f75b5', 'bg_color':'#ddebf7'})
    cfmt_Z = workbook.add_format({'color': '#000000', 'bg_color':'#8ea9db'})
    cfmt_KV = workbook.add_format({'color': '#215967', 'bg_color':'#daeef3'})
    cfmt_OS = workbook.add_format({'color': '#5b3151', 'bg_color':'#e4dfec'})
    cfmt_X = workbook.add_format({'color': '#c05065', 'bg_color':'#ff8596'})
    cfmt_TORIGHT = workbook.add_format({'color': '#963634', 'bg_color':'#fde9d9'})
    cfmt_m = workbook.add_format({'color': '#a6afc0', 'bg_color':'#d9d9d9'})
    cfmt_S = workbook.add_format({'color': '#ff7979', 'bg_color':'#ffafaf'})
    cfmt_product_bottom_line = workbook.add_format({'bottom':1, 'bottom_color':'#0010a7'})
    cfmt_sunday_col = workbook.add_format({'bg_color':'#f2f2f2', 'bold': True})
    cfmt_name_row = workbook.add_format({'bold': True})

    worksheet.set_column(0, columns_count, 30, cfmt_all)

    worksheet.set_column(0,1, 0) # Первые два столбца индекс и ин_айди
    worksheet.set_column(2,2, 60, cfmt_header_col) # Столбец с названием 
    worksheet.set_column(3,3, 2, cfmt_header_col) # Столбец цеха
    worksheet.set_column(4,13,0, cfmt_header_col) # Столбцы дат
    worksheet.set_column(first_date, columns_count, 2) # График
    
    worksheet.set_row(0, None, None, {'hidden': True})
    
    # Даты в буквы
    shop_col_num = 3 # Номер столбца shop
    dates_range_letters = '%s:%s' % (xl_rowcol_to_cell(shop_col_num, first_date), xl_rowcol_to_cell(rows_count, columns_count)) # отрезок с буквами
    shop_range_letters = '%s:%s' % (xl_rowcol_to_cell(3, shop_col_num), xl_rowcol_to_cell(rows_count, shop_col_num)) # отрезок столбца shop
    shop_range_col_letters = '%s' % xl_rowcol_to_cell(3, shop_col_num) # столбец shop
    dates_range_fin = xl_rowcol_to_cell(rows_count, columns_count)
    shop_dates_range = '%s:%s' % (xl_rowcol_to_cell(shop_col_num, 0), dates_range_fin)
    zz = '%s:%s' % (xl_rowcol_to_cell(0, first_date), dates_range_fin)
    za = '$C1'
    worksheet.conditional_format(3, first_date, rows_count, columns_count,{'type':'cell', 'criteria':'=', 'value':'"СЗ"', 'format':cfmt_SZ,})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"Ц"', 'format':cfmt_C, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"М"', 'format':cfmt_M, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"К"', 'format':cfmt_K, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"З"', 'format':cfmt_Z, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"КВ"', 'format':cfmt_KV, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"ОС"', 'format':cfmt_OS, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"X"', 'format':cfmt_X, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'">"', 'format':cfmt_TORIGHT, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"m"', 'format':cfmt_m, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(dates_range_letters, {'type':'cell', 'criteria':'=', 'value':'"S"', 'format':cfmt_S, 'multi_range': '%s %s' % (dates_range_letters, shop_range_letters)})
    worksheet.conditional_format(zz, {'type':'formula', 'criteria': '=O$3="сб"', 'format': cfmt_sunday_col})
    worksheet.conditional_format(zz, {'type':'formula', 'criteria': '=O$3="вс"', 'format': cfmt_sunday_col})
    worksheet.conditional_format(shop_dates_range, {'type':'formula', 'criteria': '=$%s="ОС"' % shop_range_col_letters , 'format': cfmt_product_bottom_line})

    worksheet.conditional_format(za, {'type':'cell', 'criteria':'=', 'value':'"Ц"', 'format':cfmt_name_row})
    
    # worksheet.set_column(20, 20, 3, cfmt_sunday_col)
    worksheet.freeze_panes(3, first_date) # Закрепление областей на странице
    # TODO
    # - Повернуть строку дат на 90
    # - Найти сегодняшнюю дату
    # - Выделить столбец сегодняшней даты
    # - Свернуть столбцы от начала до 21 день до сегодняшней даты
    # - Установить в ячейках "автоподбор ширины"
    # - Устоновить высоту строк таблицы 
    # - Высота строк всей таблицы 12
    # - Сетка для все таблицы кроме названий
    # - Отдельный лист с легендой

    worksheet.set_tab_color('#FF9900')  # Orange, цвет вкладки
    writer.save()
    return writer

if __name__ == "__main__":
    import time
    def timing():
        # Счетчик времени, таймер
        start_time = time.time()
        return lambda x: print("[{:>7.2f}с.] {}".format(time.time() - start_time, x))
    t = timing()
    df = readAppendData('testfiles\\AppendData.xlsx')
    df = colorsWrite(df)
    t("{:>5} Конец выполнения".format(''))