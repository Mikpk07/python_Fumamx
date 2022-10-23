import pymysql
import xlrd

"""
一、连接mysql数据库
"""
# 打开数据库连接
conn = pymysql.connect(
    host='localhost',  # MySQL服务器地址
    user='root',  # MySQL用户名
    password='Root1234！',  # 用户名
    charset='utf8',  # 密码
    port=3306,  # 端口
    db='test_mx',  # 数据库名称
)
print("3")

# 使用cursor()方法获取操作游标
c = conn.cursor()

"""
二、读取excel文件
"""
print("2")
FilePath = 'C:\\Users\\Administrator\\Desktop\\001.xlsx'
print("1")
# 1.打开excel文件
wkb = xlrd.open_workbook(FilePath)
# 2.获取sheet
sheet = wkb.sheet_by_index(0)  # 获取第一个sheet表['学生信息']
# 3.获取总行数
rows_number = sheet.nrows
# 4.遍历sheet表中所有行的数据，并保存至一个空列表cap[]
cap = []
for i in range(rows_number):
    x = sheet.row_values(i)  # 获取第i行的值（从0开始算起）
    cap.append(x)
print(cap)  # [['9022478', '郭赛', '男', 34.0, 'CS'], ['9022472', '林伟', '男', 36.0, 'MA'], ···]

"""
三、将读取到的数据批量插入数据库
"""
for Stu in cap:

    MXID =int(Stu[0])
    bank_received=Stu[1]
    account_name=Stu[2]
    bank_account=Stu[3]
    order_code=Stu[4]
    PO_NO=Stu[5]
    PI_NO=Stu[6]
    order_date=Stu[7]
    customer_name=Stu[8]
    order_type=Stu[9]
    send_date=Stu[10]
    pay_method=Stu[11]
    price_rule=Stu[12]
    sale_manger=Stu[13]
    currency=Stu[14]
    exchange_dollar=Stu[15]
    exchange_CNY=Stu[16]
    Manager_Account=Stu[17]
    Sales_office=Stu[18]
    Sales_group=Stu[19]
    owner=Stu[20]
    Department=Stu[21]
    close_type=Stu[22]
    close_reCOMMENTon=Stu[23]
    close_message=Stu[24]
    final_customer=Stu[25]
    company_export=Stu[26]
    sale_route=Stu[27]
    CostCenter=Stu[28]
    country=Stu[29]
    invoice_type=Stu[30]
    Shipment_status=Stu[31]
    creator=Stu[32]
    create_date=Stu[33]
    product_code=Stu[34]
    product_name=Stu[35]
    Sap_No=Stu[36]
    Sap_lineNo=Stu[37]
    num=Stu[38]
    company_price=Stu[39]
    sale_price=Stu[40]
    sale_amount=Stu[41]
    final_price=Stu[42]
    final_amount=Stu[43]
    plan_send=Stu[44]
    product_note=Stu[45]
    factory_ship=Stu[46]
    order_discount=Stu[47]
    seller_discount=Stu[48]
    return_disconut=Stu[49]
    support_discount=Stu[450]
    belong_price=Stu[51]
    belong_amount=Stu[52]
    history_factory=Stu[53]
    MRP_time=Stu[54]
    unit_nw=Stu[55]
    unit_gw=Stu[56]
    unit_volume=Stu[57]
    unit_num=Stu[58]
    close_flag=Stu[59]
    product_close_reCOMMENTon=Stu[60]
    product_close_message=Stu[61]
    real_sendnum=Stu[62]

    # 使用f-string格式化字符串，对sql进行赋值
    c.execute(f"insert into student(id, bank_received, account_name, bank_account, order_code, PO_NO, PI_NO, order_date, customer_name, order_type, send_date, pay_method, price_rule, sale_manger, currency, exchange_dollar, exchange_CNY, Manager_Account, Sales_office, Sales_group, owner, Department, close_type, close_reCOMMENTon, close_message, final_customer, company_export, sale_route, CostCenter, country, invoice_type, Shipment_status, creator, create_date, product_code, product_name, Sap_No, Sap_lineNo, num, company_price, sale_price, sale_amount, final_price, final_amount, plan_send, product_note, factory_ship, order_discount, seller_discount, return_disconut, support_discount, belong_price, belong_amount, history_factory, MRP_time, unit_nw, unit_gw, unit_volume, unit_num, close_flag, product_close_reCOMMENTon, product_close_message, real_sendnum) value ('{MXID}','{bank_received}','{account_name}','{bank_account}','{order_code}','{PO_NO}','{PI_NO}','{order_date}','{customer_name}','{order_type}','{send_date}','{pay_method}','{price_rule}','{sale_manger}','{currency}','{exchange_dollar}','{exchange_CNY}','{Manager_Account}','{Sales_office}','{Sales_group}','{owner}','{Department}','{close_type}','{close_reCOMMENTon}','{close_message}','{final_customer}','{company_export}','{sale_route}','{CostCenter}','{country}','{invoice_type}','{Shipment_status}','{creator}','{create_date}','{product_code}','{product_name}','{Sap_No}','{Sap_lineNo}','{num}','{company_price}','{sale_price}','{sale_amount}','{final_price}','{final_amount}','{plan_send}','{product_note}','{factory_ship}','{order_discount}','{seller_discount}','{return_disconut}','{support_discount}','{belong_price}','{belong_amount}','{history_factory}','{MRP_time}','{unit_nw}','{unit_gw}','{unit_volume}','{unit_num}','{close_flag}','{product_close_reCOMMENTon}','{product_close_message}','{real_sendnum}')")
conn.commit()
conn.close()
print("插入数据完成！")
