import pandas as pd

def replacesheet(path = '成本核算数据源配置表的副本.xlsx', df = pd.DataFrame()):
    with pd.ExcelWriter(
        path, 
        mode='a', 
        engine='openpyxl', 
        if_sheet_exists='replace'
    ) as writer:
        df.T.to_excel(writer, sheet_name='项目广告收入')