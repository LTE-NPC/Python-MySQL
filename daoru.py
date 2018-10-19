# excel导入到MySQL
import pymysql
import xlrd

db = pymysql.connect(host="127.0.0.1", user="yunyi", passwd="yunyi158173", db="movies", charset='utf8')


def open_excel():
    book = xlrd.open_workbook("users.xlsx")
    sheet = book.sheet_by_name("users")
    return sheet


def insert_deta():
    sheet = open_excel()
    cursor = db.cursor()
    row_num = sheet.nrows
    for i in range(1, row_num):
        row_data = sheet.row_values(i)
        value = (row_data[0], row_data[1], row_data[2], row_data[3], row_data[4])
        print(i)
        sql = "INSERT INTO tb_users(User_id,Gender,Age,Occupation,Zip_code)VALUES(%s,%s,%s,%s,%s)"
        cursor.execute(sql, value)
        db.commit()


open_excel()
insert_deta()
