
"""
Готовые заказы

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
        # logger.error("serviceNoteReadFile error with file: %s - %s" % (sn_file, ind))
        print("serviceNoteReadFile error with file: %s - %s" % (a_file, ind))
    else:
        return df

def colorsWrite(df):
    # Установка параметров для всей таблицы
    # df = df.style.set_properties(**{
    #     'font-size': '10pt',
    #     # 'datetime_format':'%d%m%Y'
    # })

    # Скрыть столбцы
    # df = df.style.hide_columns(['order_plan_start', 'work_start', 'work_end_plan', 'work_end_fact', 'zinc', 'rubberizing', 'material', 'order_plan_shipment_from', 'order_plan_shipment_before', 'order_finish'])
    
    # df = df.style.hide_index()
    # df = df.style.highlight_max(axis=0)
    # Формат даты для столбцов
    # df = df.style.set_properties(color="white", align="right")
    # df = df.style.set_properties(**{'background-color': 'yellow'})
    # def letter_color(value):
    #     if value == 'Ц':
    #         return 'color: #d6a300; background-color:#fff2c8; text-align: center'
    #     elif value == 'М':
    #         return 'color: #548235; background-color:#e2efda; text-align: center'
    #     elif value == 'К':
    #         return 'color: #2f75b5; background-color:#ddebf7; text-align: center'
    #     elif value == 'З':
    #         return 'color: #000000; background-color:#8ea9db; text-align: center'
    #     elif value == 'КВ':
    #         return 'color: #215967; background-color:#daeef3; text-align: center'
    #     elif value == 'ОС':
    #         return 'color: #5b3151; background-color:#e4dfec; text-align: center'
    #     elif value == 'X':
    #         return 'color: #c05065; background-color:#ff8596; text-align: center'
    #     elif value == '>':
    #         return 'color: #963634; background-color:#fde9d9; text-align: center'
    #     elif value == 'СЗ':
    #         return 'color: #f0800a; background-color:#fde9d9; text-align: center'
    #     elif value == 'm':
    #         return 'color: #a6afc0; background-color:#d9d9d9; text-align: center'
    #     elif value == 'S':
    #         return 'color: #ff7979; background-color:#ffafaf; text-align: center'
    #     else:
    #         return 'color:black'

    # def highlight_greaterthan_1(x):
    #     if x['shop'] == 'ОС':
    #         return ['border-bottom-width:1px; border-bottom-color:#0010a7' for v in x]
    #     else:
    #         return ['border-bottom-width:1px; border-bottom-color:#cacaca' for v in x]


    # df = df.style \
    # .applymap(letter_color) \
    # .format({'total_amt_usd_pct_diff': "{:.2%}"}) \
    # .set_properties(**{'font-size': '10pt; vertical-align: middle'}) \
    # .apply(highlight_greaterthan_1, axis=1) \
    # .hide_index()

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
    print(za)
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
    worksheet.conditional_format(zz, {'type':'formula', 'criteria': '=O$3="сб"', 'format': cfmt_sunday_col})#'=$A$1>5'
    worksheet.conditional_format(zz, {'type':'formula', 'criteria': '=O$3="вс"', 'format': cfmt_sunday_col})#'=$A$1>5'
    worksheet.conditional_format(shop_dates_range, {'type':'formula', 'criteria': '=$%s="ОС"' % shop_range_col_letters , 'format': cfmt_product_bottom_line})#'=$A$1>5'

    worksheet.conditional_format(za, {'type':'cell', 'criteria':'=', 'value':'"Ц"', 'format':cfmt_name_row})
    
    # worksheet.set_column(20, 20, 3, cfmt_sunday_col)
    worksheet.freeze_panes(3, first_date)
    worksheet.write_comment(2, 2, 'This is a comment \n New Comment', {'color': '#daeef3'}) # Комментарий 
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
    # df.to_excel("testfiles\\AppendDataColors.xlsx", engine='xlsxwriter')
    t("{:>5} Конец выполнения".format(''))