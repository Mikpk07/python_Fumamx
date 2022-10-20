import pandas as pd
from sqlalchemy import create_engine

# 初始化数据库连接，使用pymysql模块
# MySQL的用户：root, 密码:你的密码, 端口：3306,数据库：trust
engine = create_engine("mysql+pymysql://root:Root1234!@127.0.0.1:3306/hwt-yzj", encoding='utf-8')

# 查询语句，选出testexcel表中的所有数据
sql = "select * from argeeement_summary"

# read_sql_query的两个参数: sql语句， 数据库连接
df = pd.read_sql_query(sql, con=engine)

# 输出testexcel表的查询结果
#print(df)

# 创建一个writer对象, 里面的参数是一个新的表格文件名
writer = pd.ExcelWriter('test001.xlsx')
# 利用to_excel()方法将不同的数据框及其对应的sheet名称写入该writer对象中
df.to_excel(writer,sheet_name='test1')
# 数据写出到excel文件中,最后保存
writer.save()
#测试提交

