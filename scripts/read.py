import numpy as np
import pandas as pd

# 
# 1. 读取excel文件作为数据源
# 2. 计算费用
#
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
# read data source.
revenue = pd.read_excel(excel_name, sheet_name=sheet_name['source']['revenue'])
spend = pd.read_excel(excel_name, sheet_name=sheet_name['source']['spend'])
timesheet = pd.read_excel(excel_name, sheet_name=sheet_name['source']['timesheet'])
employee = pd.read_excel(excel_name, sheet_name=sheet_name['source']['employee'])
cloud = pd.read_excel(excel_name, sheet_name=sheet_name['source']['cloud'])
technical = pd.read_excel(excel_name, sheet_name=sheet_name['source']['technical'])
operation = pd.read_excel(excel_name, sheet_name=sheet_name['source']['operation'])
bjdepartment = pd.read_excel(excel_name, sheet_name=sheet_name['source']['bjdepartment'])
cddepartment = pd.read_excel(excel_name, sheet_name=sheet_name['source']['cddepartment'])
# read data config.
employee_config = pd.read_excel(excel_name, sheet_name=sheet_name['config']['employee'])
project_config = pd.read_excel(excel_name, sheet_name=sheet_name['config']['project'])

# 数据清洗
# revenue收入数据清洗
revenue = revenue.fillna(method='ffill')
revenue_adv = revenue[revenue['收入类型'] == '广告']
revenue_inapp = revenue[revenue['收入类型'] == '内购']
revenue = revenue[revenue['收入类型'] == '合计']
revenue_adv.set_index('项目收入占比',inplace=True)
revenue_inapp.set_index('项目收入占比',inplace=True)
revenue.set_index('项目收入占比',inplace=True)
revenue_percent = revenue.loc['占比'][1:]


cloud.set_index('费用项',inplace=True)
project_config.fillna(0, inplace=True)
project_config.set_index('项目', inplace=True)
# 数据运算
cloud_fee_items = cloud.index.values
cloud_project_config = project_config[cloud_fee_items]
_cloud_aliyun_fee = cloud_project_config * revenue_percent
cloud_aliyun_fee = _cloud_aliyun_fee / _cloud_aliyun_fee.sum() * cloud['阿里云']['费用']
# target fees
revenue_adv = [] 
revenue_inp = []
fee_cloud = []
fee_techinal = []
fee_dev_employee= []
fee_dev_apportion = []

fee_manage_operation = []
fee_manage_market = []
fee_manage_functional = []
fee_manage_public = []

