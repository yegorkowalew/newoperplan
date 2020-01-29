
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
            usecols = ['ID', 'Готово', 'Статус в Плане производства'],
            index_col=0
        )
    except Exception as ind:
        # logger.error("serviceNoteReadFile error with file: %s - %s" % (sn_file, ind))
        print("serviceNoteReadFile error with file: %s - %s" % (ready_file, ind))
    else:
        df = df.rename(columns={
            'ID':'inid',
            'Готово':'ready',
            'Статус в Плане производства':'status'
            }
        )
        level_map = {'-': False}
        df['produced'] = df['status'].map(level_map)
        level_map = {'+': True}
        df['ready_status'] = df['ready'].map(level_map)
        df = df.drop(['status', 'ready'], axis=1)
        df['ready_status'] = df['ready_status'].fillna(False)
        df['produced'] = df['produced'].fillna(True)
        return df

if __name__ == "__main__":
    from settings import READY_FILE
    print(readyReadFile(READY_FILE))