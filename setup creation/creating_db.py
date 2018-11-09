# Creating Data base for name retrival
import sqlite3

connection=sqlite3.connect('backlog_report.db')
#Connecting to the database
c=connection.cursor()

#c.execute ("""create table IF NOT EXISTS backlog_report (name varchar2(100))""")

op=c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='backlog_report'")
print (op)
for i,data in enumerate(op.fetchall()):
    print (i)
    print (data)
    if data[0]==1:
        print ('Table Alreay exists with some data')
        op1=c.execute("SELECT * from backlog_report")       
        for i1,data1 in enumerate(op1.fetchall()):
            print (data1)
    else:
        #creating the table if not available
        c.execute("""create table IF NOT EXISTS backlog_report (name varchar2(100))""")
        print ('table is created')
        c.execute("insert into backlog_report(name) values  ('Arivanand Murugesan')")
        c.execute("insert into backlog_report (name) values ('Jayaprakash Subramanian')") 
        c.execute("insert into backlog_report (name) values ('Krishna Nagarajan')")
        c.execute("insert into backlog_report (name) values ('Lalithkiran Gopikrishna') ")
        c.execute("insert into backlog_report (name) values ('Mohamed Musthafa Kani') ")
        c.execute("insert into backlog_report (name) values ('Praveena Mohanasundaram') ")
        c.execute("insert into backlog_report (name) values ('Ravindran Naarayanan') ")
        c.execute("insert into backlog_report (name) values ('Suchitra Chandrasekaran')") 
        c.execute("insert into backlog_report (name) values ('Suriya S Kanthan') ")
        c.execute("insert into backlog_report (name) values ('Yuvaraj Murugan') ")
        c.execute("insert into backlog_report (name) values ('Anitha Thangavel') ")
        c.execute("insert into backlog_report (name) values ('Asmita Prakash') ")
        c.execute("insert into backlog_report (name) values ('Diwakar Vasu') ")
        c.execute("insert into backlog_report (name) values ('Naresh Elango') ")
        c.execute("insert into backlog_report (name) values ('Aravind Subrmanian') ")
        c.execute("insert into backlog_report (name) values ('Santhosh Mahadevan') ")
        c.execute("insert into backlog_report (name) values ('Jeyalakshmi Sivaselvaraj')  ")
        c.execute("insert into backlog_report (name) values ('Sariga Suresh') ")
        c.execute("insert into backlog_report (name) values ('Madan Chenchuraju') ")
        c.execute("insert into backlog_report (name)  values ('Kavitha Sekar')")
        print ('Some the data are inserted')
        connection.commit()
        print ('Chanegs have been commited in DB')
op=c.execute("SELECT * from backlog_report")       
for i,data in enumerate(op.fetchall()):
    print (data)
#c.execute(""" Alter table backlog_report add path varchar2(1000) """)


def insert(Name):
    global connection
    c=connection.cursor()
    c.execute("insert into backlog_report (name) values (?)",(Name,))
    connection.commit()
def delete(Name):
    global connection
    c=connection.cursor()
    op=c.execute("delete from backlog_report where name =?",(Name,))
    print (op.fetchall())
    connection.commit()

def Paetec_open_query():
    global connection
    c=connection.cursor()
    op=c.execute('select * from backlog_report')
    petec_open1='''('Assigned Group*+' = "IT-OSS - M6 (PAETEC)") AND (NOT ('Status*' = "Resolved" OR 'Status*' = "Closed" OR 'Status*' = "Cancelled")) AND ( 'Case Type*' ="Incident" ) AND  '''
    text=''' 'Assignee+' = "'''
    petec_open2=''
    for num,i in enumerate(op.fetchall()):
        #print (i[0])
        if num ==0:
            petec_open2=petec_open2+text+i[0]+'''"'''
        else:
            petec_open2=petec_open2+' OR '+text+i[0]+'''"'''        
    #print (petec_open1 + '( '+petec_open2+' )' )
    return petec_open1 + '( '+petec_open2+' )' 
    connection.commit()
    
def Paetec_Closed_query(current_month,next_month,respective_year1,respective_year2):
    global connection
    c=connection.cursor()
    op=c.execute('select * from backlog_report')
    petec_Closed1='''('Assigned Group*+' = "IT-OSS - M6 (PAETEC)")  AND ( 'Case Type*' ="Incident" ) AND '''
    petec_Closed3=''' AND (('Last Resolved Date' >= "'''+str(current_month)+'/01/'+str(respective_year1)+'''" AND 'Last Resolved Date' < "'''+str(next_month)+'/01/'+str(respective_year2)+'"))'
    text=''' 'Assignee+' = "'''
    petec_Closed2=''
    for num,i in enumerate(op.fetchall()):
        #print (i[0])
        if num ==0:
            petec_Closed2=petec_Closed2+text+i[0]+'''"'''
        else:
            petec_Closed2=petec_Closed2+' OR '+text+i[0]+'''"'''        
    #print (petec_Closed1 + '( '+petec_Closed2+' )' + petec_Closed3)
    return petec_Closed1 + '( '+petec_Closed2+' )' + petec_Closed3
    connection.commit()
def paetec_request_query():
    global connection
    c=connection.cursor()
    op=c.execute('select * from backlog_report')
    paetec_request1='''('Assigned Group*+' = "IT-OSS - M6 (PAETEC)") AND (NOT ('Status*' = "Resolved" OR 'Status*' = "Closed" OR 'Status*' = "Cancelled")) AND  '''
    text=''' 'Assignee+' = "'''
    paetec_request2=''
    for num,i in enumerate(op.fetchall()):
        #print (i[0])
        if num ==0:
           paetec_request2=paetec_request2+text+i[0]+'''"'''
        else:
           paetec_request2=paetec_request2+' OR '+text+i[0]+'''"'''        
    #print (paetec_request1 + '( '+paetec_request2+' )'  + ''' AND ( 'Case Type*' ="Request" )''')
    return paetec_request1 + '( '+paetec_request2+' )' + ''' AND ( 'Case Type*' ="Request" ) '''
    connection.commit()
def paetec_ris_query():
    global connection
    c=connection.cursor()
    op=c.execute('select * from backlog_report')
    paetec_ris1='''('Assigned Group*+' = "IT-OSS - M6 (PAETEC)") AND (NOT ('Status*' = "Resolved" OR 'Status*' = "Closed" OR 'Status*' = "Cancelled")) AND  '''
    text=''' 'Assignee+' = "'''
    paetec_ris2=''
    for num,i in enumerate(op.fetchall()):
        #print (i[0])
        if num ==0:
           paetec_ris2=paetec_ris2+text+i[0]+'''"'''
        else:
           paetec_ris2=paetec_ris2+' OR '+text+i[0]+'''"'''        
    #print (paetec_ris1 + '( '+paetec_ris2+' )'  + ''' AND ( 'Case Type*' ="Ris" )''')
    return paetec_ris1 + '( '+paetec_ris2+' )' + ''' AND ( 'Case Type*' ="Ris" ) '''
    connection.commit()
def close_connection():
    global connection
    connection.close()
    print ('DB connection has been disabled')
    quit()
#Paetec_open_query()
