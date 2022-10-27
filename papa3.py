from tkinter.ttk import Combobox
import pyodbc
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

inputdb = 0
outdb = 1
lst1=[]
lst2=[]
fieldname1=[]
fieldname2=[]
tablenames1=[]
tablenames2=[]
tab1=1
tab2=2

def but1():
    global inputdb,tablenames1
    filepath = filedialog.askopenfilename()
    filepath=filepath.replace('\\','/')
    inputdb = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+filepath+';')
    tabcur1=inputdb.cursor()
    temptablenames1=tabcur1.tables()
    for i in temptablenames1:
        if i[3]=='TABLE':
            tablenames1.append(i[2])
    cb1.configure(values=tablenames1)
    b1.configure(bg='green')
    
def but2():
    global outdb
    filepath = filedialog.askopenfilename()
    filepath=filepath.replace('\\','/')
    outdb = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+filepath+';')
    tabcur2=outdb.cursor()
    temptablenames2=tabcur2.tables()
    for i in temptablenames2:
        if i[3]=='TABLE':
            tablenames2.append(i[2])
    cb2.configure(values=tablenames2)
    b2.configure(bg='green')
    
def everything1():
    global fieldname2,fieldname1,tab1,tab2
    tab1=str(cb1.get())
    tab2=str(cb2.get())
    namecur1 = inputdb.cursor()
    namecur1.execute('select * from '+tab1)
    fieldname1 = [column[0] for column in namecur1.description]
    namecur2 = outdb.cursor()
    namecur2.execute('select * from '+tab2)
    fieldname2 = [column[0] for column in namecur2.description]
    
def everything2():
    global lst1,lst2,tab1,tab2
    cursor1 = inputdb.cursor()
    cursor2 = outdb.cursor()
    if len(lst1)==1:
        cursor1.execute('select '+str(lst1[0])+' from '+tab1)
        r1=cursor1.fetchall()
        for rec in r1:
            cursor2.execute('insert into '+tab2+' ('+str(lst2[0])+') values '+str(rec).replace(",",""))
        messagebox.showinfo("Succesfully Completed","Records inserted Succesfully !!")
    elif len(lst1)==0:
        messagebox.showerror("No Connections Made","Please Form Links Before Proceeding")
    else:
        cursor1.execute('select '+(", ".join(lst1))+' from '+tab1)
        r1=cursor1.fetchall()
        for rec in r1:
            cursor2.execute('insert into '+tab2+' ('+(", ".join(lst2))+') values '+str(rec))
        messagebox.showinfo("Succesfully Completed","Records inserted Succesfully !!")
    outdb.commit()
    
def connect():
    global lst1,lst2
    lst1.append(str(c21.get()))
    lst2.append(str(c22.get()))
    list21.insert("end",str(c21.get())+" --> "+str(c22.get()))
    
def menu():
    global fieldname2,fieldname1,c21,c22,list21
    try:
        everything1()
    except:
        messagebox.showerror("Fatal Error","Error Occured While Trying To Open Database !! Exiting ...")
        exit(0)
    men=Toplevel(root)
    men.title("Make Connections")
    l21=Label(men,text="Source File").grid(row=1,column=1,sticky="news",padx=10,pady=10)
    l22=Label(men,text="Destination File").grid(row=1,column=2,sticky="news",padx=10,pady=10)
    c21=Combobox(men)
    c21['values']=tuple(fieldname1)
    c21.grid(row=2,column=1,sticky="news",padx=10,pady=10)
    c22=Combobox(men)
    c22['values']=tuple(fieldname2)
    c22.grid(row=2,column=2,sticky="news",padx=10,pady=10)
    b21=Button(men,text="Create Link",command=connect)
    b21.grid(row=3,column=1,columnspan=2,sticky="news",padx=10,pady=10)
    list21=Listbox(men,height=10,width=10)
    list21.grid(row=4,column=1,columnspan=2,sticky="news",padx=10,pady=10)
    b22=Button(men,text="Process",command=everything2)
    b22.grid(row=5,column=1,columnspan=2,sticky="news",padx=10,pady=10)
    men.mainloop()
    
root=Tk()
root.title("Choose Tables")
l1=Label(text="Source File").grid(row=1,column=1,sticky="news",padx=10,pady=10)
l2=Label(text="Destination File").grid(row=3,column=1,sticky="news",padx=10,pady=10)
b1=Button(text="Choose",command=but1)
b1.grid(row=1,column=2,sticky="news",padx=10,pady=10)
b2=Button(text="Choose",command=but2)
b2.grid(row=3,column=2,sticky="news",padx=10,pady=10)
b3=Button(text="Next",command=menu)
b3.grid(row=5,column=1,columnspan=2,sticky="news",padx=10,pady=10)
cb1=Combobox(root)
cb1['values']=tuple(tablenames1)
cb1.grid(row=2,column=1,columnspan=2,padx=10,pady=10,sticky="news")
cb2=Combobox(root)
cb2['values']=tuple(tablenames2)
cb2.grid(row=4,column=1,columnspan=2,padx=10,pady=10,sticky="news")
root.mainloop()