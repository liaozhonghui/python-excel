# python-excel数据分析工具

背景： excel处理

用途：使用

应用：日常报表自动化， 日报，周报，月报
数据分析师

#### 基础环境搭建
anacoda: xlwings + numpy + pandas + matplolib
xlwings: python的图表处理工具
numpy: 提供通用的数值数据处理的计算基础, 适合处理统一的数值数组数据
pandas: 专门为处理表格和混杂数据设计（Series和DataFrame）
matplotlib: 图形绘制工具

#### 数据源、数据清洗
pandas使用浮点值NaN(Not a Number)表示缺失数据

1. 处理缺失数据 （过滤和填充）
dropna 根据标签中的值是否存在缺失数据进行过滤
fillna 用指定值或者插值方法(ffill或者bfill)填充缺失数据
isnull 返回一个有布尔值的对象，这些布尔值表示哪些值是缺失值Na
notnull isnull的否定式

2. 数据转换（去重， map转化， replace替换）
duplicated 检查是否是重复行
drop_duplicated 删除重复行
使用map是一种实现元素级转换以及其他数据清理工作的便捷方式。

3. 字符串操作

#### 数据处理
1. 数据的聚合与分组
2. 数据运算
3. 时间序列
4. 数据分组、数据透视表
5. 多表连接
6. 数据可视化

#### 报表自动化
日报，周报，月报


#### 参考书籍：
《对比Excel，轻松学习Python数据分析》
《利用Python进行数据分析》（第二版）