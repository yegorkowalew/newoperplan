
"""
Готовые заказы

"""
import pandas as pd

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
            usecols = [
                'ID', 
                'Готово', 
                'Статус в Плане производства',
                'Дата',
                'Отсутствие тех документации',
                'Дефицит материалов',
                'Дефицит мощностей',
                'Нет технологической возможности (аутсорс)',
                ],
            index_col=0
        )
    except Exception as ind:
        # logger.error("serviceNoteReadFile error with file: %s - %s" % (sn_file, ind))
        print("serviceNoteReadFile error with file: %s - %s" % (ready_file, ind))
    else:
        df = df.rename(columns={
            'ID':'in_id',
            'Готово':'ready',
            'Статус в Плане производства':'status',
            'Дата':'ready_date',
            'Отсутствие тех документации':'failure_1',
            'Дефицит материалов':'failure_2',
            'Дефицит мощностей':'failure_3',
            'Нет технологической возможности (аутсорс)':'failure_4',
            }
        )
        level_map = {'-': False}
        df['produced'] = df['status'].map(level_map)
        level_map = {'+': True}
        df['ready_status'] = df['ready'].map(level_map)
        df = df.drop(['status', 'ready'], axis=1)
        df['ready_status'] = df['ready_status'].fillna(False)
        df['produced'] = df['produced'].fillna(True)

        df['ready_date'] = pd.to_datetime(df['ready_date'], errors='coerce')
        df = df[[
            'produced',
            'ready_status',
            'ready_date',
            'failure_1',
            'failure_2',
            'failure_3',
            'failure_4',
        ]]
        return df

if __name__ == "__main__":
    from settings import READY_FILE
    readyDf = readyReadFile(READY_FILE)
    readyDf.to_excel("testfiles\\Ready.xlsx")