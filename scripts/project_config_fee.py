import pandas as pd
import numpy as np

excel_name = '成本核算数据源配置表.xlsx'
sheet_name = {
    'source': {
        'revenue': '收入',
        'spend':'投放费',
        'timesheet':'项目工时',
        'employee':'人员工资奖金五险一金',
        'cloud':'云服务费',
        'technical':'技术服务费',
        'operation':'运营费用',
        'bjdepartment':'北京天天部门费用',
        'cddepartment':'成都天天部门费用',
    },
    'config': {
        'employee':'人员费用配置表',
        'project':'项目费用配置表'
    }
}

revenue = pd.read_excel(excel_name, sheet_name=sheet_name['source']['revenue'])
revenue

revenue = revenue.fillna(method='ffill')
revenue_adv = revenue[revenue['收入类型'] == '广告']
revenue_inapp = revenue[revenue['收入类型'] == '内购']
revenue = revenue[revenue['收入类型'] == '合计']
revenue_adv.set_index('项目收入占比',inplace=True)
revenue_inapp.set_index('项目收入占比',inplace=True)
revenue.set_index('项目收入占比',inplace=True)
revenue.drop('收入类型', 1, inplace=True)
revenue

revenue_percent = revenue.loc['占比'][1:]
revenue_percent_df = pd.DataFrame(revenue_percent)
revenue_percent_df

cloud = pd.read_excel(excel_name, sheet_name=sheet_name['source']['cloud'])
cloud.set_index('费用项',inplace=True)
cloud

project_config = pd.read_excel(excel_name, sheet_name=sheet_name['config']['project'])
project_config.columns = list(project_config.columns.values[0:4]) + list(project_config.loc[0][4:].values)
project_config = project_config.drop(0).fillna(0)
project_config.set_index('项目', inplace= True)
project_config


project_config_fee_percent = pd.merge(project_config,revenue_percent_df, left_index=True, right_index=True, how='left')
project_config_fee_percent
print('数据清洗完毕')
print('开启数据计算')


for col in project_config_fee_percent.columns.values[3:-1]:
    fee_percent = project_config_fee_percent[col] * project_config_fee_percent['占比']
    project_config_fee_percent[col] = fee_percent / fee_percent.sum() * 100

project_config_fee_percent


project_fee = project_config_fee_percent.copy()

for col in project_config_fee_percent.columns.values[3:-1]: 
    if col in cloud.index:
        project_fee[col] = project_config_fee_percent[col] * cloud.loc[col, '费用'] / 100
        
project_fee


technical = pd.read_excel(excel_name, sheet_name=sheet_name['source']['technical'])
technical.set_index('费用项',inplace=True)
technical_nomatched = technical.loc['不匹配项目技术服务费']
technical.drop('不匹配项目技术服务费', inplace=True)

operation = pd.read_excel(excel_name, sheet_name=sheet_name['source']['operation'])
operation.set_index('费用项',inplace=True)
operation_nomatched_bj = operation.loc['北京天天不匹配项目运营费']
operation.drop('北京天天不匹配项目运营费', inplace=True)
operation_nomatched_cd = operation.loc['成都天天不匹配项目运营费']
operation.drop('成都天天不匹配项目运营费', inplace=True)

project_fee['技术服务费匹配项目费用'] = 0
project_fee['运营费用匹配项目费用'] = 0
project_fee

for p in technical.index:
    if p in project_fee.index:
        project_fee['技术服务费匹配项目费用'][p] = technical.loc[p]['费用']
        
project_fee

for p in operation.index:
    if p in project_fee.index:
        project_fee['运营费用匹配项目费用'][p] = operation.loc[p]['费用']
        
project_fee

from sqlalchemy import create_engine
engine = create_engine('postgresql://postgres:postgres@localhost:5432/finance')
engine

py_project_fee = project_fee.convert_dtypes()
py_project_fee.columns = ['project_group','dirproject', 'main',
'bj-aliyun-cloud', 'bj-google-cloud', 'cd-aliyun-cloud', 'cd-googleandaws-cloud', 'cd-otherproject-cloud', 
'technical', 'bj-operation', 'cd-operation',
'percent', 'not-matched-technical', 'not-matched-operation']
py_project_fee.index.name = 'project'
py_revenue = revenue.T.convert_dtypes()
py_revenue.columns = ['revenue', 'percent']
py_revenue.index.name = 'project'
py_cloud = cloud.convert_dtypes()
py_cloud.columns = ['fee']
py_cloud.index.name='fee-item'

py_technical = pd.read_excel(excel_name, sheet_name=sheet_name['source']['technical'])
py_operation = pd.read_excel(excel_name, sheet_name=sheet_name['source']['operation'])
py_technical = py_technical.convert_dtypes()
py_technical.columns = ['fee-item','fee']
py_operation = py_operation.convert_dtypes()
py_operation.columns = ['fee-item','fee']

with engine.connect() as conn:
    tran = conn.begin()
    try: 
        py_project_fee.to_sql('test_project_fee',conn, if_exists = 'append')
        py_revenue.to_sql('test_revenue',conn, if_exists = 'append')
        py_cloud.to_sql('test_cloud', conn, if_exists = 'append')
        py_technical.to_sql('test_technical', conn, if_exists = 'append', index = False)
        py_operation.to_sql('test_operation', conn, if_exists = 'append', index = False)
        tran.commit()
    except:
        print('code, rollback')
        tran.rollback()

py_operation

