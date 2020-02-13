
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
            return 'color: #d6a300; background-color:#fff2c8; text-align: center'
        elif value == 'М':
            return 'color: #548235; background-color:#e2efda; text-align: center'
        elif value == 'К':
            return 'color: #2f75b5; background-color:#ddebf7; text-align: center'
        elif value == 'З':
            return 'color: #000000; background-color:#8ea9db; text-align: center'
        elif value == 'КВ':
            return 'color: #215967; background-color:#daeef3; text-align: center'
        elif value == 'ОС':
            return 'color: #5b3151; background-color:#e4dfec; text-align: center'
        elif value == 'X':
            return 'color: #c05065; background-color:#ff8596; text-align: center'
        elif value == '>':
            return 'color: #963634; background-color:#fde9d9; text-align: center'
        elif value == 'СЗ':
            return 'color: #f0800a; background-color:#fde9d9; text-align: center'
        elif value == 'm':
            return 'color: #a6afc0; background-color:#d9d9d9; text-align: center'
        elif value == 'S':
            return 'color: #ff7979; background-color:#ffafaf; text-align: center'
        else:
            return 'color:black'

    def highlight_greaterthan_1(x):
        if x['shop'] == 'ОС':
            return ['border-bottom-width:1px; border-bottom-color:#0010a7' for v in x]
        else:
            return ['border-bottom-width:1px; border-bottom-color:#cacaca' for v in x]


    df = df.style \
    .applymap(letter_color) \
    .format({'total_amt_usd_pct_diff': "{:.2%}"}) \
    .set_properties(**{'font-size': '10pt; vertical-align: middle'}) \
    .apply(highlight_greaterthan_1, axis=1) \
    .hide_index()
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