'''import numpy
data =numpy.loadtxt("G:\python folder\preferences")
headers = [] 
for c in range(1,data.shape[0]): 
    if data[c, 0] != "": 
        headers.append(data[c, 0])
print(headers)'''
import xlrd
import pyodbc
conn=pyodbc.connect("DRIVER={SQL Server};server=localhost;database=placement_data")
conn1=pyodbc.connect("DRIVER={SQL Server};server=localhost;database=placement_data")
wb= xlrd.open_workbook("G:\python folder\IT HOD\preferences.xlsx")
sheet=wb.sheet_by_index(0)
attributes=[]
modified_roll=''
for i in range(sheet.ncols):
    if sheet.cell_value(0,i) is not '':
        attributes.append(sheet.cell_value(0,i))
length=len(attributes)

if length==5:
    for i in range(1,sheet.nrows):
        row_obj=[]
        for j in range(sheet.ncols):
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
        if len(row_obj)==5:
            conn.execute("insert into preferences(Name,student_id,Branch,preference_1,preference_2) values(?,?,?,?,?)",(row_obj[0],row_obj[1],row_obj[2],row_obj[3],row_obj[4]))
            conn.commit()
            del row_obj  


if length==6:
    for i in range(1,sheet.nrows):
        row_obj=[]
        for j in range(sheet.ncols):
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
                if(sheet.cell_value(i,j)=='NA' or sheet.cell_value(i,j)=='NILL'):
                    row_obj.append('-')
                else:
                    row_obj.append(sheet.cell_value(i,j))
        if len(row_obj)==6:
            print(row_obj)
            conn.execute("insert into preferences(Name,student_id,Branch,preference_1,preference_2,preference_3) values(?,?,?,?,?,?)",(row_obj[0],row_obj[1],row_obj[2],row_obj[3],row_obj[4],row_obj[5]))
            conn1.execute("insert into prefer_numbers values(?)",(row_obj[1]))
            conn.commit()
            conn1.commit()
            del row_obj  


if length==7:
    for i in range(1,sheet.nrows):
        row_obj=[]
        for j in range(sheet.ncols):
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
        if len(row_obj)==7:
            conn.execute("insert into preferences(Name,student_id,Branch,preference_1,preference_2,preference_3,preference_4) values(?,?,?,?,?,?,?)",(row_obj[0],row_obj[1],row_obj[2],row_obj[3],row_obj[4],row_obj[5],row_obj[6]))
            conn.commit()
            del row_obj  

if length==8:
    for i in range(1,sheet.nrows):
        row_obj=[]
        for j in range(sheet.ncols):
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
        if len(row_obj)==8:
            conn.execute("insert into preferences(Name,student_id,Branch,preference_1,preference_2,preference_3,preference_4,preference_5) values(?,?,?,?,?,?,?,?)",(row_obj[0],row_obj[1],row_obj[2],row_obj[3],row_obj[4],row_obj[5],row_obj[6],row_obj[7]))
            conn.commit()
            del row_obj  



if length==9:
    for i in range(1,sheet.nrows):
        row_obj=[]
        for j in range(sheet.ncols):
            if j==1:
                if('-' in str(sheet.cell_value(i,j))):
                    del row_obj
                    row_obj=[]
                    break
            if sheet.cell_value(i,j) is not '':
                row_obj.append(sheet.cell_value(i,j))
        if len(row_obj)==9:
            conn.execute("insert into preferences(Name,student_id,Branch,preference_1,preference_2,preference_3,preference_4,preference_5,preference_6) values(?,?,?,?,?,?,?,?,?)",(row_obj[0],row_obj[1],row_obj[2],row_obj[3],row_obj[4],row_obj[5],row_obj[6],row_obj[7],row_obj[8]))
            conn.commit()
            del row_obj  

"""

if length==10:
    for i in range(1,sheet.nrows):
        row_obj=[]
        for j in range(sheet.ncols):
            if j==1:
                if('-' in sheet.cell_value(i,j)):
                    continue
            if sheet.cell_value(i,j) is not '':
                row_obj.apeend(sheet.cell_value(i,j))
        conn.execute("insert into preferences(Name,student_id,Branch,preference_1,preference_2) values(?,?,?,?,?)",(row_obj[0],row_obj[1],row_obj[2],row_obj[3],row_obj[4]))
        del row_obj  



if length==5:
    for i in range(1,sheet.nrows):
        row_obj=[]
        for j in range(sheet.ncols):
            if j==1:
                if('-' in sheet.cell_value(i,j)):
                    continue
            if sheet.cell_value(i,j) is not '':
                row_obj.apeend(sheet.cell_value(i,j))
        conn.execute("insert into preferences(Name,student_id,Branch,preference_1,preference_2) values(?,?,?,?,?)",(row_obj[0],row_obj[1],row_obj[2],row_obj[3],row_obj[4]))
        del row_obj
"""
data=[]





conn.close()


        
