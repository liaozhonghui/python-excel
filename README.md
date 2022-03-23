# python-excel 分享

### 使用python处理excel的方式
xlwings + numpy + pandas + matplotlib
优势：
1. 日常报表： 日报、周报、月报

分析指标：
概览指标
对比型指标：同比、环比

数据分析的常规流程：
熟悉工具-> 明确目的 -> 获取数据 -> 熟悉数据 -> 处理数据 -> 分析数据 -> 得出结论 -> 展示结论
处理数据：
- 异常数据
- 重复数据
- 缺失数据
- 测试数据

### python 的相关模块
pandas
1. series数据结构
```py
import pandas as pd
s1 = pd.Series(['a', 'b', 'c', 'd'])
```
2. 使用index方法获取Series的索引
3. 使用values方法获取Series的值

DataFrame表格型结构
```py
import pandas as pd
df = pd.DataFrame(['a', 'b', 'c', 'd'])
```

数据源读取
- pd.read_excel(path string, sheet_name='sheet1', header = 0, usecols = [0, 2]) header表示使用第几行作为列索引，usecols表示使用哪些列
- pd.read_csv(path string, sep = " ", nrows = 2) sep表示分隔符号, nrows表示使用前面多少行
- pd.read_table() 类似pd.read_csv, 默认分隔符为\t
- 使用import postgres

### 数据清洗
dataframe
- dropna 根据各标签的值中是否存在缺失数据对轴标签进行过滤，可以通过阈值调节对缺失值的容忍度
- fillna 使用指定值或者插值方法(如ffill 或者bfill)填充缺失数据
- isnull 返回一个含有布尔值的对象
- notnull 表示isnull的否定
- drop_duplicates 去除重复值
- replace([A, B], C)表示将A、B替换成C, replace({A: a, B: b, C: c})多对多替换
- insert(2, 'colum', [datalist]) 插入新数据
- .T 使df进行转置
- melt进行宽表变长表
- apply函数 和applymap函数进行lambda运算

groupby进行数据分组，然后和数据透视表进行结合使用
pd.pivot_table(data, values = None, index=None, columns=None, aggfunc='mean', fill_value=None, margins=False, dropna=True, margins_name='All')