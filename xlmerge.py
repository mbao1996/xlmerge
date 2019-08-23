#coding=gbk
from openpyxl import load_workbook
import os

work_catalog = "c:\\PythonWork\\test\\"
fn = work_catalog + "demo.xlsx"
sheet = 'Sheet1'
opr = []

row_start = 3
col_opr = 7

def add_str_to_list(list, str):
    str_exist = False
    for i in range(len(list)):
        print('===', i)
        if( list[i] == str ):
            str_exist = True
            break
    if( str_exist == False ):
        list.append(str)

try:
    wb = load_workbook(fn,keep_vba=True)
except Exception as e:
    print(str(e))
    os._exit(0)
ws = wb[sheet]

print(ws.max_row)
for i in range(row_start, ws.max_row+1):
    op = ws.cell(i, col_opr).value
    op = ''.join(op.split())
    print(i, '---', ws.cell(i, 1).value, '[', ws.cell(i, col_opr).value, ']')
    add_str_to_list(opr, op)
    
print(opr)