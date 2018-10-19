import pymysql
import xlwt


def export_excel(table_name):
    conn = pymysql.connect(host='localhost', port=3306,
                          db='test1', user='root',
                          passwd='123456', charset='utf8')
    # 建立游标
    cur = conn.cursor()
    sql = 'select * from test1'
    # 执行mysql
    cur.execute(sql)
    # 列表生成式，所有字段
    fileds = [filed[0] for filed in cur.description]
    # 所有数据
    all_data = cur.fetchall()
    # 写excel
    # 先创建一个book
    book = xlwt.Workbook()
    # 创建一个sheet表
    sheet = book.add_sheet('sheet1')
    # enumerate自动计算下标
    # 跟上面的代码功能一样
    for col, field in enumerate(fileds):
        sheet.write(0, col, field)

    # 从第一行开始写
    # 行数
    row = 1
    # 二维数据，有多少条数据，控制行数
    for data in all_data:
        # 控制列数
        for col, field in enumerate(data):
            sheet.write(row, col, field)
        # 每次写完一行，行数加1
        row += 1
    # 保存excel文件
    book.save('%s.xls' % table_name)


export_excel('app_student')


