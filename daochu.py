# mysql导出到excel

import xlwt
import pymysql

conn = pymysql.connect(host="127.0.0.1", user="root", passwd="123456", db="test1", charset='utf8')

cursor = conn.cursor()

count = cursor.execute("select * from test1")

print(count)

cursor.scroll(0, mode='absolute')

results = cursor.fetchall()
print('=======',results)
fields = cursor.description
print('*****',fields)
wbk = xlwt.Workbook()
print('*&&&&&&&&',wbk)
sheet = wbk.add_sheet('tb_movie', cell_overwrite_ok=True)
print('@@@@@@@@',sheet)
for i in range(0, len(fields)):
    sheet.write(0, i, fields[i][0])

ics = 1
jcs = 0
for ics in range(1, len(results) + 1):
    for jcs in range(0, len(fields)):
        sheet.write(ics, jcs, results[ics-1][jcs])

wbk.save('movie.xlsx')
print('()(())()()())(()()()()')
