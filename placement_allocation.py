import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
"""
student_info=pd.read_excel("Total.xlsx")
Capgemini=pd.read_excel("Capgemini Results 10.09.2018.xlsx")
Infosys=pd.read_excel("Final selects-Infosys Ltd.xlsx")
CTS=pd.read_excel("CTS 2019 Batch CBIT Selects.xlsx")
preferences=pd.read_excel("Day-1 Sharing Preferences.xlsx")""""
import cgi, os
import cgitb; cgitb.enable()
form = cgi.FieldStorage()
fileitem = form['master']
fileitem1 = form['preference']
fileitem2 = form['infosys']
fileitem3 = form['ctg']
fileitem4 = form['cpg']


if fileitem.filename:
    fn = os.path.basename(fileitem.filename)
    student_info=pd.read_excel(fn)
else:
    message = 'No file was uploaded'
if fileitem4.filename:
    fn = os.path.basename(fileitem4.filename)
    Capgemini=pd.read_excel(fn)
else:
    message = 'No file was uploaded'
if fileitem3.filename:
    fn = os.path.basename(fileitem3.filename)
    CTS=pd.read_excel(fn)
else:
    message = 'No file was uploaded'
if fileitem2.filename:
    fn = os.path.basename(fileitem2.filename)
    Infosys=pd.read_excel(fn)
else:
    message = 'No file was uploaded'
if fileitem1.filename:
    fn = os.path.basename(fileitem1.filename)
    preferences=pd.read_excel(fn)
else:
    message = 'No file was uploaded'

student_info.set_index('Roll.No.',inplace=True)
Capgemini.set_index('Roll Number',inplace=True)
Infosys.set_index('Roll No.',inplace=True)
CTS.set_index('Current University Reg No',inplace=True)
preferences.set_index('Roll No.',inplace=True)
for roll in student_info.index:
    if type(roll)== str and '-' in roll:
        roll_split=roll.split('-')
        modified_roll=''
        for i in roll_split:
            modified_roll+=i
        modified_roll=int(modified_roll)
        student_info=student_info.rename({roll: modified_roll})

final_Infosys=[]
final_Capgemini=[]
final_CTS=[]

preferences=preferences.groupby(level=0).first()

first_preference = (preferences.index)
##converting '1601-##-###-###' to 1601########
for roll in first_preference:
    if type(roll)== str and '-' in roll:
        roll_split=roll.split('-')
        modified_roll=''
        for i in roll_split:
            modified_roll+=i
        modified_roll=int(modified_roll)
        preferences=preferences.rename({roll: modified_roll})

first_preference = preferences.index

for roll in first_preference:
    if preferences['Preference-1'][roll] == 'Infosys':
        if roll in Infosys.index:
            final_Infosys.append(roll)
    elif preferences['Preference-1'][roll] == 'Cognizant':
        if roll in CTS.index:
            final_CTS.append(roll)
    elif preferences['Preference-1'][roll] == 'Capgemini':
        if roll in Capgemini.index:
            final_Capgemini.append(roll)

second_preference = [roll for roll in first_preference if roll not in final_CTS and roll not in final_Capgemini and roll not in final_Infosys]

for roll in second_preference:
    if preferences['Preference-2'][roll] == 'Infosys':
        if roll in Infosys.index:
            final_Infosys.append(roll)
    elif preferences['Preference-2'][roll] == 'Cognizant':
        if roll in CTS.index:
            final_CTS.append(roll)
    elif preferences['Preference-2'][roll] == 'Capgemini':
        if roll in Capgemini.index:
            final_Capgemini.append(roll)

third_preference=[roll for roll in second_preference if roll not in final_CTS and roll not in final_Capgemini and roll not in final_Infosys]

for roll in third_preference:
    if preferences['Preference-3'][roll] == 'Infosys':
        if roll in Infosys.index:
            final_Infosys.append(roll)
    elif preferences['Preference-3'][roll] == 'Cognizant':
        if roll in CTS.index:
            final_CTS.append(roll)
    elif preferences['Preference-3'][roll] == 'Capgemini':
        if roll in Capgemini.index:
            final_Capgemini.append(roll)

not_placed=[roll for roll in third_preference if roll not in final_CTS and roll not in final_Capgemini and roll not in final_Infosys]
import xlwt 
from xlwt import Workbook 
  
# Workbook is created 
wb = Workbook() 
excel_name=input('enter file name')
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet 1')
student_info['Company selected']=pd.np.nan
i=0
cell_format02.set_num_format('###0')
for roll in student_info.index:
    
    sheet1.write(i,0,str(roll),cell_fotmat2)
    i=i+1
    if roll in final_CTS:
        student_info['Company selected'][roll]='Cognizant'
    elif roll in final_Capgemini:
        student_info['Company selected'][roll]='Capgemini'
    elif roll in final_Infosys:
        student_info['Company selected'][roll]='Infosys'
    elif roll in not_placed:
        student_info['Company selected'][roll]='Not placed'
    else:
        student_info['Company selected'][roll]='Not participated'
    sheet1.write(i-1,1,str(student_info['Company selected'][roll]))
#print(student_info['Company selected'])
wb.save('companies1.xls')

















"""
import numpy as np
import pandas as pd

student_info=pd.read_excel("Total.xlsx")
cap=pd.read_excel("Capgemini Results 10.09.2018.xlsx")
infy=pd.read_excel("Final selects-Infosys Ltd.xlsx")
cts=pd.read_excel("CTS 2019 Batch CBIT Selects.xlsx")
preference=pd.read_excel("Day-1 Sharing Preferences.xlsx")

cap.set_index(cap.columns[1],inplace=True)
cts.set_index(cts.columns[0],inplace=True)
infy.set_index(infy.columns[0],inplace=True)
preference.set_index(preference.columns[1],inplace=True)
student_info.set_index(student_info.columns[1],inplace=True)

a=['Infosys','Capgemini','Cognizant']

all3=[]
inf_cts=[]
cts_cap=[]
cap_inf=[]
oinf=[]
ocap=[]
octs=[]
for roll in preference.index.values:
        if roll in infy.index.values :
            if roll in cts.index.values:
                if roll in cap.index.values:
                    all3.append(roll)
                else:
                    inf_cts.append(roll)
        if roll in infy.index.values :
            if roll in cap.index.values:
                if roll in cts.index.values:
                    pass
                else:
                    cap_inf.append(roll)
        if roll in cap.index.values :
            if roll in cts.index.values:
                if roll in infy.index.values:
                    pass
                else:
                    cts_cap.append(roll)
        if roll in infy.index.values :
            if roll not in cts.index.values:
                if roll not in cap.index.values:
                    oinf.append(roll)
        if roll in cts.index.values :
            if roll not in infy.index.values:
                if roll not in cap.index.values:
                    octs.append(roll)
        if roll in cap.index.values :
            if roll not in cts.index.values:
                if roll not in infy.index.values:
                    ocap.append(roll)
        
print(len(all3),
len(inf_cts),
len(cts_cap),
len(cap_inf),
len(oinf),
len(ocap),
len(octs))
Infosys=oinf
Capgemini=ocap
Cognizant=octs

for roll in all3:
    if preference.loc[roll]['Preference-1']==a[0]:
        Infosys.append(roll)
    elif preference.loc[roll]['Preference-1']==a[1]:
        Capgemini.append(roll)
    elif preference.loc[roll]['Preference-1']==a[2]:
        Cognizant.append(roll)
for roll in inf_cts:
    if preference.loc[roll]['Preference-1']==a[0]:
        Infosys.append(roll)
    elif preference.loc[roll]['Preference-1']==a[2]:
        Cognizant.append(roll)
    elif preference.loc[roll]['Preference-2']==a[0]:
        Infosys.append(roll)
    elif preference.loc[roll]['Preference-2']==a[2]:
        Cognizant.append(roll)
for roll in cts_cap:
    if preference.loc[roll]['Preference-1']==a[2]:
        Cognizant.append(roll)
    elif preference.loc[roll]['Preference-1']==a[1]:
        Capgemini.append(roll)
    elif preference.loc[roll]['Preference-2']==a[2]:
        Cognizant.append(roll)
    elif preference.loc[roll]['Preference-2']==a[1]:
        Capgemini.append(roll)
for roll in cap_inf:
    if preference.loc[roll]['Preference-1']==a[0]:
        Infosys.append(roll)
    elif preference.loc[roll]['Preference-1']==a[1]:
        Capgemini.append(roll)
    elif preference.loc[roll]['Preference-2']==a[0]:
        Infosys.append(roll)
    elif preference.loc[roll]['Preference-2']==a[1]:
        Capgemini.append(roll)
        
print(len(Capgemini),len(Infosys),len(Cognizant))
"""
