import xlrd
import pyodbc
conn=pyodbc.connect("DRIVER={SQL Server};server=localhost;database=placement_data")
wb= xlrd.open_workbook("G:\python folder\IT HOD\Cognizont.xlsx")
sheet=wb.sheet_by_index(0)

x=input("enter the company name\n")

for i in range(1,sheet.nrows):
    if('-' in str(sheet.cell_value(i,0))):
        continue
    elif sheet.cell_value(i,1) is not '':
        conn.execute("insert into {} values({})".format(x,sheet.cell_value(i,0)))
        conn.commit()
        print(sheet.cell_value(i,0)) 
