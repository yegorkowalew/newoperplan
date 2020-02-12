
"""
Готовые заказы

"""
import pandas as pd

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
    def letter_color(value):
        if value == 'Ц':
            return 'color: #d6a300; background-color:#fff2c8'
        elif value == 'М':
            return 'color: #548235; background-color:#e2efda'
        elif value == 'К':
            return 'color: #2f75b5; background-color:#ddebf7'
        elif value == 'З':
            return 'color: #000000; background-color:#8ea9db'
        elif value == 'КВ':
            return 'color: #215967; background-color:#daeef3'
        elif value == 'ОС':
            return 'color: #5b3151; background-color:#e4dfec'
        elif value == 'X':
            return 'color: #c05065; background-color:#ff8596'
        elif value == '>':
            return 'color: #963634; background-color:#fde9d9'
        elif value == 'СЗ':
            return 'color: #f0800a; background-color:#fde9d9'
        elif value == 'm':
            return 'color: #a6afc0; background-color:#d9d9d9'
        elif value == 'S':
            return 'color: #ff7979; background-color:#ffafaf'
        else:
            return 'color:black'
        # elif value == 'М':
        #     color = 'green'
        # elif value == 'К':
        #     color = 'black'

        # return 'color: %s' % color



    df = df.style \
    .applymap(letter_color, subset=[303, 304, 305, 306, 307, 308, 309, 310, 311, 312, 313, 314, 315, 316, 317, 318, 319, 320, 321, 322, 323, 324, 325, 326, 327, 328, 329, 330, 331, 332, 333, 334, 335, 336, 337, 338, 339, 340, 341, 342, 343, 344, 345, 346, 347, 348, 349, 350, 351, 352, 353, 354, 355, 356, 357, 358, 359, 360, 361, 362, 363, 364, 365, 366, 367, 368, 369, 370, 371, 372, 373, 374, 375, 376, 377, 378, 379, 380, 381, 382, 383, 384, 385, 386, 387, 388, 389, 390, 391, 392, 393, 394, 395, 396, 397, 398, 399, 400, 401, 402, 403, 404, 405, 406, 407, 408, 409, 410, 411, 412, 413, 414, 415, 416, 417, 418, 419]) \
    .format({'total_amt_usd_pct_diff': "{:.2%}"})
    return df

if __name__ == "__main__":
    import time
    def timing():
        # Счетчик времени, таймер
        start_time = time.time()
        return lambda x: print("[{:>7.2f}с.] {}".format(time.time() - start_time, x))
    t = timing()
    df = readAppendData('testfiles\\AppendData.xlsx')
    df = colorsWrite(df)
    df.to_excel("testfiles\\AppendDataColors.xlsx", engine='openpyxl')
    t("{:>5} Конец выполнения".format(''))