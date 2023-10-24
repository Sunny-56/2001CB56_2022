
from datetime import datetime
from unicodedata import name
start_time = datetime.now()

def attendance_report():
###Code
 import csv
 import os
 import numpy as np
 os.system('cls')
 no_r_l=[]
 nme_std=[]

 with open('input_registered_students.csv', 'r') as file:
  reader = csv.reader(file)
  num=0
  for row in reader:
   if num!=0:
    no_r_l.append(row[0])
    nme_std.append(row[1])
   num=num+1 
  num=num-1
 dat =["28/07","01/08","04/08","08/08","11/08","15/08","18/08","22/08","25/08","29/08","01/09","05/09","08/09","12/09","15/09","26/09","29/09"]
 std=len(dat)


 with open('input_attendance.csv', 'r') as file:
  reader = csv.reader(file) 

  lst_m=[]
  lst_c=[] 
  lst_d=[]  
  for x in no_r_l:
   lst_c=[]
   for j in dat:
    lst_d=[]
    a=0
    b=0
    e=0
    with open('input_attendance.csv', 'r') as file:
     reader = csv.reader(file)
     for row in reader:
      if j==row[0][0:5] and row[1][0:8]==x and (row[0][11:13]=="14" or row[0][11:16]=="15:00" ) and a==0 : 
       a=a+1
      
      elif j==row[0][0:5] and row[1][0:8]==x :
        e=e+1
      elif j==row[0][0:5] and row[1][0:8]==x and (row[0][11:13]=="14" or row[0][11:16]=="15:00") and a>0 : 
        b=b+1
     lst_d.append(a+b+e)  
     lst_d.append(a)
     lst_d.append(b)
     lst_d.append(e)
     if a==0:
      lst_d.append(1)
     else:
      lst_d.append(0) 
    lst_c.append(lst_d)
    lst_m.append(lst_c)
 
 
 if os.path.exists("output"):
  for f in os.listdir("output"):
    os.remove(os.path.join("output",f))

 os.chdir("output")
 from openpyxl import Workbook
 for i in range(0,num): 
  book=Workbook()
  sheet= book.active    
  all_r_w=[] 
  all_r_w.append(["Date","Roll No.","Name","Total attendance count","Real","Duplicate","Invalid","absent"])
  all_r_w.append(["",no_r_l[i],nme_std[i],"","","","",""])
  for q in range(0,std):
   all_r_w.append([dat[q],"","",lst_m[i][q][0],lst_m[i][q][1],lst_m[i][q][2],lst_m[i][q][3],lst_m[i][q][4]])
  for w in all_r_w:
   sheet.append(w)
  book.save( no_r_l[i] + ".xlsx")

  dc_t={0:"A",1:"P"}
  book=Workbook()
  sheet= book.active    
  all_r_w=[]
  rl_nm=["Roll No.","Name"]
  for i in dat:
   rl_nm.append(i)
  rl_nm.append("Total Lecture taken")
  rl_nm.append("Total Real")
  rl_nm.append("% Attendance")
  all_r_w.append(rl_nm) 

  rl_nm=["(Sorted by roll no.)","","Atleast one real P"]
  for i in range(0,std-1):
   rl_nm.append("")
  rl_nm.append("(=Total Mon+Thur dynamic count")
  rl_nm.append("")
  rl_nm.append("Real/Actual Lecture taken")
  all_r_w.append(rl_nm) 

  for i in range(0,num): 
   rl_nm=[no_r_l[i],nme_std[i]]
   c=0
   for q in range(0,std):
    rl_nm.append(dc_t[lst_m[i][q][1]])
    if dc_t[lst_m[i][q][1]]=="P":
     c=c+1
   rl_nm.append(std)
   rl_nm.append(c)
   l=(c/std)*100
   rl_nm.append("{:.2f}".format(l))
   all_r_w.append(rl_nm)

  for w in all_r_w:
   sheet.append(w)
  book.save( "Attendance_report_consolidated" + ".xlsx") 

attendance_report()


#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))



