
"""
План производства

"""
import pandas as pd

def productionPlanReadFile(input_file):
    try:
        df = pd.read_excel(
            input_file, 
            sheet_name="Даты",
            header=0,
            index_col=0,
            usecols = [
                'ID',
                '104 Дата начала',
                '104 Дата окончания',
                '104 Дата цинкования',
                '102 Дата начала',
                '102 Дата окончания',
                '102 Дата цинкования',
                '101 Дата начала',
                '101 Дата окончания',
                '101 Дата цинкования',
                '107 Дата начала',
                '107 Дата окончания',
                '107 Дата цинкования',
                ],
        )
    except Exception as ind:
        # logger.error("serviceNoteReadFile error with file: %s - %s" % (sn_file, ind))
        print("productionPlanReadFile error with file: %s - %s" % (input_file, ind))
    else:
        df = df.rename(columns={
            'ID':'in_id',
            '104 Дата начала':'104_start_date',
            '104 Дата окончания':'104_end_date',
            '104 Дата цинкования':'104_zinc',
            '102 Дата начала':'102_start_date',
            '102 Дата окончания':'102_end_date',
            '102 Дата цинкования':'102_zinc',
            '101 Дата начала':'101_start_date',
            '101 Дата окончания':'101_end_date',
            '101 Дата цинкования':'101_zinc',
            '107 Дата начала':'107_start_date',
            '107 Дата окончания':'107_end_date',
            '107 Дата цинкования':'107_zinc',
            }
        )
        return df

if __name__ == "__main__":
    from settings import PRODUCTION_PLAN_FILE
    productionPlanDf = productionPlanReadFile(PRODUCTION_PLAN_FILE)
    productionPlanDf.to_excel("testfiles\\Production_Plan.xlsx")