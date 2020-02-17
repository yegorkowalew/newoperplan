import itertools
from readshedule import sheduleWorker
from settings import READY_FILE, SN_FILE, IN_DOCUMENT_FILE, IN_DOCUMENT_FOLDER, PRODUCTION_PLAN_FILE, SHEDULE_FOLDER
import time
import os
import pandas as pd

def sheduleReadFile(file_path):
    try:
        df = pd.read_excel(
            file_path, 
            sheet_name="График",
            header=5,
            usecols = [
                'Узел',
                'Наименование',
                'Операция',
                'Кол-во\nУзел',
                'Кол-во\nИзделие',
                'Кол-во\nЗаказ',
                'Кол-во\nМассив',
                'Материал',
                'Размер',
                'Размер заготовки',
                'Росцеховка',
                ],
        )
    except Exception as ind:
        # logger.error("serviceNoteReadFile error with file: %s - %s" % (sn_file, ind))
        print("serviceNoteReadFile error with file: %s - %s" % (file_path, ind))
    else:
        df = df.rename(columns={
            'Узел':'node',
            'Наименование':'name',
            'Операция':'operation',
            'Кол-во\nУзел':'qtynode',
            'Кол-во\nИзделие':'qtyproduct',
            'Кол-во\nЗаказ':'qtyorder',
            'Кол-во\nМассив':'qtyarray',
            'Материал':'material',
            'Размер':'size',
            'Размер заготовки':'workpiecesize',
            'Росцеховка':'shops',
        })
        return df

def mergeSheduleFiles():
    base_path = 'C:\\work\\newoperplan\\testfiles'
    file_path = '2311894 - ТСЦ-320Ц-35м.xlsx'
    order_no = '2311894'
    dispatcher = 'Пирлик Інна Іванівна'
    shedule_file_path = os.path.join(base_path, file_path)
    
    df = sheduleReadFile(shedule_file_path)
    nn = itertools.count(start=0, step=1)
    # def append_node_num(node_val):
    #     return next(nn)
    df.at[0, 'node_num'] = ''
    old_node = ''
    old_node_num = next(nn)
    for index, row in df.iterrows():
        if df.at[index, 'node'] == old_node:
            df.at[index, 'node_num'] = old_node_num
        elif not pd.isnull(df.at[index, 'node']):
            # print(df.at[index, 'node'] == old_node)
            old_node = df.at[index, 'node']
            df.at[index, 'node_num'] = old_node_num
            old_node_num = next(nn)
        # df['node_num']  
    # df['dispatcher'] = dispatcher
    print(df.head(50))
    # print('------------')
    # print(df['node'])

if __name__ == "__main__":
    def timing():
        # Счетчик времени, таймер
        start_time = time.time()
        return lambda x: print("[{:>7.2f}с.] {}".format(time.time() - start_time, x))
    # df = pd.DataFrame({
    #     'dispatcher':['Сосяк Наталія Олексіївна', 'Мудренко Наталія Володимирівна', 'Хрупало Інна Василівна', 'Пирлик Інна Іванівна'],
    #     'file_path':[
    #         'C:\\work\\newoperplan\\testfiles\\Графики ПДО\\Диспетчер-1 - Сосяк Наталія Олексіївна\\20 - Агро-Рось\\231700014 - самоплив А1-ДСП-50.49.000-10.xlsx',
    #         'C:\\work\\newoperplan\\testfiles\\Графики ПДО\\Диспетчер-2 - Мудренко Наталія Володимирівна\\124 - Старокостянтинівський олійноекстракційний завод 15 шт\\2311715 - СМВУ.110.10.К65.В12 + Логотип KMZ INDUSTRIES.xlsx',
    #         'C:\\work\\newoperplan\\testfiles\\Графики ПДО\\Диспетчер-3 - Хрупало Інна Василівна\\5 - ТОВ Амбар Експорт БКВ\\231700007 - Узел подшипниковый МШЗ.220313.000 Н.xlsx',
    #         'C:\\work\\newoperplan\\testfiles\\Графики ПДО\\Диспетчер-4 - Пирлик Інна Іванівна\\Задел\\2317650 - Направляюча.xlsx'
    #     ],
    #     'order_no':[
    #         '',
    #     ],
    #     'date_creation':[],
    #     'date_modification':[],
    # })

    t = timing()

    # correct_files, error_files = sheduleWorker(SHEDULE_FOLDER)
    # print(correct_files)
    # correct_files.to_excel('testfiles\\Correct_Shedule_Files.xlsx')
    # error_files.to_excel('testfiles\\Failure_Shedule_Files.xlsx')
    mergeSheduleFiles()
    t("{:>5} Обработка графиков".format(''))