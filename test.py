# -*- coding: UTF-8 -*-
import pyodbc
 
# 连接数据库（不需要配置数据源）,connect()函数创建并返回一个 Connection 对象
cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=c:\\PythonWork\\test\\test.accdb')
# cursor()使用该连接创建（并返回）一个游标或类游标的对象
crsr = cnxn.cursor()
 
# 打印数据库goods.mdb中的所有表的表名
print('`````````````` goods ``````````````')
for table_info in crsr.tables(tableType='TABLE'):
    print(table_info.table_name)
 
for row in crsr.execute("SELECT * from 测试表1"):
    print(row)

crsr.execute("select * from 测试表1  ORDER BY key desc")
# 每列的详细信息
des = crsr.description
# 获取表头
print("表头:", ",___".join([item[0] for item in des]))

crsr.commit()
crsr.close()
cnxn.close()