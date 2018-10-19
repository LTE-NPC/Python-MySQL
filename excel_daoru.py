# 使用python把excel导入到MySQL
import xlrd
import pymysql

# 打开数据所在的路径表名
book = xlrd.open_workbook('test2.xls')
# 这个是表里的sheet名称
sheet = book.sheet_by_name('sheet')

# 建立一个 MySQL连接
conn = pymysql.connect(
    host='localhost',
    user='root',
    passwd='123456',
    db='test1',
    port=3306,
    charset='utf8'
)

# 获得游标
cur = conn.cursor()

# 创建插入sql语句
query = 'insert into test2(brach,perform_brach,' \
        'brach_id,profes_bracha,' \
        'error,money)values(%s,%s,%s,%s,%s,%s)'

# 创建一个for循环迭代读取xls文件每行数据的，
# 从第二行开始是要跳过标题行
# 括号里面1表示从第二行开始(计算机是从0开始数)
for r in range(1, sheet.nrows):
    # (r, 0)表示第二行的0就是表里的A1:A1
    brach = sheet.cell(r, 0).value
    perform_brach = sheet.cell(r, 1).value
    brach_id = sheet.cell(r, 2).value
    profes_bracha = sheet.cell(r, 3).value
    error = sheet.cell(r, 4).value
    money = sheet.cell(r, 5).value
    values = (brach, perform_brach, brach_id, profes_bracha, error, money)
    # 执行sql语句
    cur.execute(query, values)

# close关闭文档
cur.close()
# commit 提交
conn.commit()
# 关闭MySQL链接
conn.close()
# 显示导入多少列
columns = str(sheet.ncols)
# 显示导入多少行
rows = str(sheet.nrows)
print('导入'+columns+'列'+rows+'行数据到MySQL数据库!')
