#coding=gbk
from openpyxl import load_workbook
import os

work_catalog = "c:\\PythonWork\\test\\"
fn = work_catalog + "demo.xlsm"
sheet = 'Sheet1'
opr = []

row_start = 3
col_opr = 7
col_chg =[8,9,10]

def add_str_to_list(list, str):
    str_exist = False
    for i in range(len(list)):
        if( list[i] == str ):
            str_exist = True
            break
    if( str_exist == False ):
        list.append(str)
def open_fn(fn, wr):
    try:
        if( wr == True ):
            wb = load_workbook(fn,keep_vba=True)
        else:
            wb = load_workbook(fn)
    except Exception as e:
        print(str(e))
        os._exit(0)
    return(wb)
def save_fn(wb, fn):
    try:
        wb.save(fn)
    except Exception as e:
        print(str(e))
        os._exit(0)

wb = open_fn(fn, True)
ws = wb[sheet]

# get opr
for i in range(row_start, ws.max_row+1):
    op = ws.cell(i, col_opr).value
    op = ''.join(op.split())
    add_str_to_list(opr, op)
for k in range(len(opr)):   # every files to be merged
    print('-----------', opr[k])
    fn_div = fn.split('.')
    fn_rd = fn_div[0] + '_' + opr[k] + '.' + fn_div[1]

    wb_rd = open_fn(fn_rd, False)
    ws_rd = wb_rd[sheet]
    rd = None
    rd_sn = 0
    wr = None
    wr_sn = 0
    for i in range(row_start, ws_rd.max_row+1):     # 副文件的每一行
        if( ws_rd.cell(i, 1).value != None ):
            rd = ws_rd.cell(i, 1).value
            rd_sn = 0
        else:
            rd_sn += 1
        if( ws_rd.cell(i, col_opr).value == opr[k] ):   # 确认副文件的操作者
            for j in range(row_start, ws.max_row+1):    # 主文件中搜索
                if( ws.cell(j, 1).value != None ):
                    wr = ws.cell(j, 1).value
                    wr_sn = 0
                else:
                    wr_sn += 1
                if( ws.cell(j, col_opr).value == opr[k] ):   # 主文件中确认副文件的操作者
                    if( wr == rd and wr_sn == rd_sn ):      # 确认是同一行
                        for col in range(len(col_chg)):
#                            print(rd, ' * ', i,'-',j, '---', col_chg[col])
#                            print(ws_rd.cell(i, col_chg[col]).value)
                            ws.cell(j, col_chg[col]).value = ws_rd.cell(i, col_chg[col]).value
                        break

save_fn(wb, fn)

print("work done")
