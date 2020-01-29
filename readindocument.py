
"""
График документации

"""
import pandas as pd
import os

def inDocumentReadFile(path, file_name):
    try:
        df = pd.read_excel(
            os.path.join(path, file_name),
            sheet_name="График",
            header=3,
            usecols = [
                'ID',
                'ДПКД % 1',
                'ДПКД Дата 1',
                'ДПКД Дата 2',
                'ДПКД Дата 3',
                'ДПОД % 1',
                'ДПОД Дата 1',
                'ДПОД Дата 2',
                'ДПОД Дата 3',
                'ДПКОД % 1',
                'ДПКОД Дата 1',
                'ДПКОД Дата 2',
                'ДПКОД Дата 3',
                'ДПИЧ % 1',
                'ДПИЧ Дата 1',
                'ДПИЧ Дата 2',
                'ДПИЧ Дата 3'
            ],

            parse_dates = [
                'ДПКД Дата 1',
                'ДПКД Дата 2',
                'ДПКД Дата 3',
                'ДПОД Дата 1',
                'ДПОД Дата 2',
                'ДПОД Дата 3',
                'ДПКОД Дата 1',
                'ДПКОД Дата 2',
                'ДПКОД Дата 3',
                'ДПИЧ Дата 1',
                'ДПИЧ Дата 2',
                'ДПИЧ Дата 3'
            ],

            dtype={
                'ID':int,
            },
            # converters={
            #     'Готово':bool,
            #     'Комплектовочные Выдача':bool,
            #     'Отгрузочные Выдача':bool,
            #     'Конструкторские Выдача':bool,
            #     'Изменения чертежей Выдача':bool,
            # }
        )

    except Exception as ind:
        # logger.error("inDocumentReadFile error with file: %s - %s" % (file_name, ind))
        print("inDocumentReadFile error with file: %s - %s" % (file_name, ind))
    else:
        df = df.rename(columns={
            'ID':'in_id',
            'Диспетчер':'dispatcher',
            'ДПКД % 1':'pickup_issue',
            'ДПКД Дата 1':'pickup_date_1',
            'ДПКД Дата 2':'pickup_date_2',
            'ДПКД Дата 3':'pickup_date_3',
            'ДПОД % 1':'shipping_issue',
            'ДПОД Дата 1':'shipping_date_1',
            'ДПОД Дата 2':'shipping_date_2',
            'ДПОД Дата 3':'shipping_date_3',
            'ДПКОД % 1':'design_issue',
            'ДПКОД Дата 1':'design_date_1',
            'ДПКОД Дата 2':'design_date_2',
            'ДПКОД Дата 3':'design_date_3',
            'ДПИЧ % 1':'drawings_issue',
            'ДПИЧ Дата 1':'drawings_date_1',
            'ДПИЧ Дата 2':'drawings_date_2',
            'ДПИЧ Дата 3':'drawings_date_3'
            }
        )
        df['dispatcher'] = os.path.split(path)[-1]
        df['pickup_date_1'] = pd.to_datetime(df['pickup_date_1'])
        df['pickup_date_2'] = pd.to_datetime(df['pickup_date_2'])
        df['pickup_date_3'] = pd.to_datetime(df['pickup_date_3'])
        df['shipping_date_1'] = pd.to_datetime(df['shipping_date_1'])
        df['shipping_date_2'] = pd.to_datetime(df['shipping_date_2'])
        df['shipping_date_3'] = pd.to_datetime(df['shipping_date_3'])
        df['design_date_1'] = pd.to_datetime(df['design_date_1'])
        df['design_date_2'] = pd.to_datetime(df['design_date_2'])
        df['design_date_3'] = pd.to_datetime(df['design_date_3'])
        df['drawings_date_1'] = pd.to_datetime(df['drawings_date_1'])
        df['drawings_date_2'] = pd.to_datetime(df['drawings_date_2'])
        df['drawings_date_3'] = pd.to_datetime(df['drawings_date_3'])

        df = df.astype({
            # 'pickup_issue':'object',
            'pickup_date_1':'object',
            'pickup_date_2':'object',
            'pickup_date_3':'object',
            'shipping_date_1':'object',
            'shipping_date_2':'object',
            'shipping_date_3':'object',
            'design_date_1':'object',
            'design_date_2':'object',
            'design_date_3':'object',
            'drawings_date_1':'object',
            'drawings_date_2':'object',
            'drawings_date_3':'object',
            })
        def setup_pickup_date(row):
            return max([row['pickup_date_1'], row['pickup_date_2'], row['pickup_date_3']])
        
        def setup_shipping_date(row):
            return max([row['shipping_date_1'], row['shipping_date_2'], row['shipping_date_3']])

        def setup_design_date(row):
            return max([row['design_date_1'], row['design_date_2'], row['design_date_3']])

        def setup_drawings_date(row):
            return max([row['drawings_date_1'], row['drawings_date_2'], row['drawings_date_3']])

        df['pickup_date'] = df.apply (lambda row: setup_pickup_date(row), axis=1)
        df['shipping_date'] = df.apply (lambda row: setup_shipping_date(row), axis=1)
        df['design_date'] = df.apply (lambda row: setup_design_date(row), axis=1)
        df['drawings_date'] = df.apply (lambda row: setup_drawings_date(row), axis=1)

        # level_map = {0: False}
        # df['pickup_issue_f'] = df['pickup_issue'].map(level_map)
        # df['pickup_issue_f'] = df['pickup_issue_f'].fillna(False)
        def setup_pickup_issue(row):
            if row['pickup_issue'] == 0:
                return False
            else:
                return True
        df['pickup_issue_f'] = df.apply (lambda row: setup_pickup_issue(row), axis=1)

        def setup_shipping_issue(row):
            if row['shipping_issue'] == 0:
                return False
            else:
                return True
        df['shipping_issue_f'] = df.apply (lambda row: setup_shipping_issue(row), axis=1)

        def setup_design_issue(row):
            if row['design_issue'] == 0:
                return False
            else:
                return True
        df['design_issue_f'] = df.apply (lambda row: setup_design_issue(row), axis=1)

        def setup_drawings_issue(row):
            if row['drawings_issue'] == 0:
                return False
            else:
                return True
        df['drawings_issue_f'] = df.apply (lambda row: setup_drawings_issue(row), axis=1)

        df = df[['in_id', 'dispatcher', 'pickup_date', 'shipping_date', 'design_date', 'drawings_date', 'pickup_issue_f', 'shipping_issue_f', 'design_issue_f', 'drawings_issue_f']]
        df = df.set_index('in_id')
        return df

def inDocumentFindFile(in_folder, need_file):
    tree = os.walk(in_folder)
    folder_list = []
    for i in tree:
        for address, dirs, files in [i]:
            for fl in files:
                if fl == need_file:
                    folder_list.append(address)
    return folder_list

def inDocumentRebuild(in_folder, need_file):
    for folder in inDocumentFindFile(in_folder, need_file):
        df = inDocumentReadFile(folder, need_file)
        df.to_excel(folder+'\\'+"output1.xlsx")
        # print(folder)
        print(df)

if __name__ == "__main__":
    from settings import IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER
    print(inDocumentRebuild(IN_DOCUMENT_FOLDER, IN_DOCUMENT_FILE))