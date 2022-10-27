import pyodbc

inputdb = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\parth\Desktop\Database4.accdb;')
namecur1 = inputdb.cursor()
cursor1 = inputdb.cursor()
namecur1.execute('select * from Source')
fieldname1 = [column[0] for column in namecur1.description]
outdb = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\parth\Desktop\Database31.accdb;')
namecur2 = outdb.cursor()
cursor2 = outdb.cursor()
namecur2.execute('select * from DEST')
fieldname2 = [column[0] for column in namecur2.description]
l1=["SName","FName","MName"]
l2=["NM","GN","MN"]
cursor1.execute('select '+(", ".join(l1))+' from Source')
r1=cursor1.fetchall()
for rec in r1:
    print('insert into DEST ('+(", ".join(l2))+') values '+str(rec))
    cursor2.execute('insert into DEST ('+(", ".join(l2))+') values '+str(rec))
outdb.commit()