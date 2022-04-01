import numpy as np
import pandas as pd

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

# 读取数据源
# 1. 项目配置表数据
# 2. 收入数据
# 
project_config = pd.read_excel(excel_name, sheet_name=sheet_name['config']['project'])
revenue = pd.read_excel(excel_name, sheet_name=sheet_name['source']['revenue'])

# 数据清洗
# 1. 清洗收入数据
revenue = revenue.fillna(method='ffill')
revenue_adv = revenue[revenue['收入类型'] == '广告']
revenue_inapp = revenue[revenue['收入类型'] == '内购']
revenue = revenue[revenue['收入类型'] == '合计']
revenue_adv.set_index('项目收入占比',inplace=True)
revenue_inapp.set_index('项目收入占比',inplace=True)
revenue.set_index('项目收入占比',inplace=True)
revenue_percent = revenue.loc['占比'][1:]
revenue_percent_df = pd.DataFrame(revenue_percent)
revenue_percent_df
# 2. 清洗project_config表数据
project_config.columns = list(project_config.columns.values[0:4]) + list(project_config.loc[0][4:].values)
project_config = project_config.drop(0).fillna(0)
project_config.set_index('项目', inplace= True)
project_config_fee_percent = pd.merge(project_config,revenue_percent_df, left_index=True, right_index=True, how='left')
for col in project_config_fee_percent.columns.values[3:-1]:
    fee_percent = project_config_fee_percent[col] * project_config_fee_percent['占比']
    project_config_fee_percent[col] = fee_percent / fee_percent.sum() * 100

project_fee = project_config_fee_percent.copy()

# 3. 清洗云服务费数据
cloud = pd.read_excel(excel_name, sheet_name=sheet_name['source']['cloud'])
cloud.set_index('费用项',inplace=True)

# 4. 清洗技术服务费数据
technical = pd.read_excel(excel_name, sheet_name=sheet_name['source']['technical'])
technical.set_index('费用项',inplace=True)
technical_nomatched = technical.loc['不能匹配项目费用']
technical.drop('不能匹配项目费用')

# 5. 清洗运营费用数据
operation = pd.read_excel(excel_name, sheet_name=sheet_name['source']['operation'])
operation.set_index('费用项',inplace=True)
operation_nomatched = operation.loc['不能匹配项目费用']
operation.drop('不能匹配项目费用')

# 数据计算
# 1. 计算云服务费
# 2. 计算技术服务费
# 3. 计算运营费用
for col in project_config_fee_percent.columns.values[3:-1]: 
    if col in cloud.index:
        project_fee[col] = project_config_fee_percent[col] * cloud.loc[col, '费用'] / 100
    elif col in ['技术服务费']:
          project_fee[col] = project_config_fee_percent[col] * technical_nomatched['费用'] / 100 
    elif col in ['运营费用']:
          project_fee[col] = project_config_fee_percent[col] * operation_nomatched.loc[col, '费用'] / 100

project_config

print(project_fee)