#encoding=utf-8
import mysql.connector
conn = mysql.connector.connect(
    host='127.0.0.1'
    # 连接名称，默认127.0.0.1
    , user='root'
    # 用户名
    , passwd='Root1234!'
    # 密码
    , port=3306
    # 端口，默认为3306
    , db='test_mx'
    # 数据库名称
    , charset='utf8'
    # 字符编码
)
cur = conn.cursor()  # 生成游标对象
sql = "select * from mx_canatueorder"  # SQL语句
cur.execute(sql)  # 执行SQL语句
data = cur.fetchall()  # 通过fetchall方法获得数据
for i in data[:10]:  # 打印输出前2条数据
    print(i)
cur.close()  # 关闭游标
print("SUCCESSS")
conn.close()  # 关闭连接
