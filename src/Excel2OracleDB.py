import os
import re
import sys
import numpy as np
import pandas as pd
import cx_Oracle
import tkinter as tk
from tkinter import *
import traceback
from datetime import datetime
#cx_Oracle.init_oracle_client(lib_dir=r"C:\oracle\instantclient_18_5_64bit")

main_path   = os.environ['USERPROFILE']
pwd_path    = "" #main_path+'\\Conf\\Pwd\\OracleDB.txt'
tab_name    = ""

module      = "Excel2OracleDB"
vers        = "v1.0.0"
root        = Tk()
root.geometry("700x400")
root.title(module+"("+vers+")")

def noww(genre):
    noww = datetime.now()
    noww = noww.strftime("%H:%M:%S")
    T.insert(END, genre+" : The Time is "+ noww+"\n")   

def impData():
    try:
        source   = data_source_t.get()
        tab_name = tab_name_t.get()        
        noww("Start")  
        col      = getRow(int(schema_t.get()))

        user     = col[1]
        password = col[2]
        host     = col[3]
        port     = col[4]
        dbname   = col[5]

        con = cx_Oracle.connect(user=user, password=password, dsn=host+':'+port+'/'+dbname)
        cur = con.cursor()
        xl = pd.ExcelFile(source)
        sh = list(xl.sheet_names)
        bol = 10000
        T.insert(END,"tabs : "+str(sh)+"\n")
        cur.arraysize = 10000
        #rowsize=180000
        #conversion for the letters of the Turkish alphabet those are not in the English alphabet
        l1=[["İ","I"],["ı","I"],["Ç","C"],["ç","C"],["Ö","O"],["ö","O"],
            ["Ü","U"],["ü","U"],["ğ","G"],["Ğ","G"],["Ş","S"],["ş","S"],
            [" ","_"]]
        l2=[["int64","INT"],["object","VARCHAR2(4000)"],["float64","NUMBER"]]
        col=[]
        dty=[]
        crt=""
        for i in sh:
            df = pd.read_excel(source, sheet_name=i, comment='#')
            T.insert(END, "the tab been processed : "+i+"...\n")
            dt = df.head(1)
            for d in df.columns:
                e = 0
                c = d                        
                for i in l1:                
                    c = c.replace(l1[e][0],l1[e][1])
                    e+= 1
                c = re.sub(r'[^a-zA-Z0-9_]', '', c)
                col.append(c)
                df[d]=df[d].fillna('').astype('str')

            for t in dt.dtypes:
                e = 0
                c = str(t)
                for k in l2:
                    c = c.replace(l2[e][0],l2[e][1])
                    e+= 1                
                dty.append(str(c))

            k=0
            for i in dty:
                crt += ''.join(str(col[k]))+'   '+str(i)+','
                k+= 1

            cur.execute('SELECT COUNT(*) FROM user_tables WHERE table_name = UPPER(:1) ',[tab_name])        
            var = cur.fetchone()[0]

            if int(var) == 0:
                crt = 'CREATE TABLE '+tab_name+' ( '+crt.rstrip(",")+' )'  
                cur.execute(crt)

            bvr=""
            for i in range(0,df.columns.size):
                bvr+= ':'+str(i+1)+','             

            sql = 'INSERT INTO '+tab_name+' VALUES( '+bvr.rstrip(",")+' )'
            for j in range(0,len(df)):
                if j%bol == 0:
                    dl = df[:bol].values.tolist()
                    df = df[bol:]
                    cur.executemany(sql,dl)
                    con.commit()
                    
    except Exception as e:        
        T.insert(END, str(e)+"\n")    

    noww("Bitis")
    cur.close()
    con.close()

Label(root, text="Connection ID / File :",fg="darkblue").pack()
schema_t = StringVar()
schema_e = Entry(root, width=2, textvariable=schema_t, bg="light gray").pack()

conn_t = StringVar()
conn_e = Entry(root, width=80, textvariable=conn_t, bg="light gray").pack()
conn_t.set(main_path)

def getRow(no):
    try:
        pwd_path = conn_t.get()
        satir = open(pwd_path,'r').readlines()
        return satir[no-1].split()

    except FileNotFoundError:
        T.insert(END, "Connection reference file cannot be found\n") 

Label(root, text="Data source file :",fg="darkblue").pack()
data_source_t = StringVar()
data_source   = Entry(root, width=80, textvariable=data_source_t, bg="light cyan").pack()
data_source_t.set(main_path)

Label(root, text="Table name :",fg="darkblue").pack()
tab_name_t = StringVar()
tab_name_e = Entry(root, width=80, textvariable=tab_name_t, bg="light cyan").pack()

Button(root,text="Import Data",width=10,command=impData,bg="darkblue",fg="white").pack()
S = Scrollbar(root)
T = Text(root,height = 14,width = 100,bg="light gray")
S.pack(side=RIGHT, fill=Y)
T.pack(side=LEFT, fill=Y)
S.config(command=T.yview)
T.config(yscrollcommand=S.set)

root.mainloop()
