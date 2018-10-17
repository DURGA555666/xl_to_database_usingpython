import xlrd
import pyodbc # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
from xlwt import easyxf
from xlutils.copy import copy
conn=pyodbc.connect("DRIVER={SQL Server};server=localhost;database=placement_data")
conn1=pyodbc.connect("DRIVER={SQL Server};server=localhost;database=placement_data")
wb= xlrd.open_workbook("G:\python folder\IT HOD\master.xlsx")
sheet=wb.sheet_by_index(0)
modified_roll=''
for i in range(1,sheet.nrows):
    row_obj=[]
    for j in range(1,sheet.ncols):
        if j==1:
            if('-' in str(sheet.cell_value(i,j))):
                roll=str(sheet.cell_value(i,j))
                roll_split=roll.split('-')
                for c in roll_split:
                    modified_roll+=c
                modified_roll=int(modified_roll)
        if sheet.cell_value(i,j) is not '':
            if modified_roll is not '':
                row_obj.append(modified_roll)
                modified_roll=''
                continue
            if(sheet.cell_value(i,j)=='NA'):
                row_obj.append('-')
            else:
                row_obj.append(sheet.cell_value(i,j))
    if len(row_obj)==9:
        print(row_obj)
        conn.execute("insert into master values(?,?,?,?,?,?,?,?,?)",(row_obj[0],row_obj[1],row_obj[2],row_obj[3],row_obj[4],row_obj[5],row_obj[6],row_obj[7],row_obj[8]))
        conn.commit()
        del row_obj
        




conn.close()


        
