#============ Importing the required libraries==========================================================
import tkinter
from  tkinter import *
from tkinter import messagebox
import creating_db
import os
import Backlog_report_generator_loop
#from hover import HoverInfo
import re
import copy_paste
import time
import final_flow
#================= Root window initialisation=====================================================
root=Tk()
root.title('Backlog Report Generation')
#root.geometry('600*400')
root.config(background="#E0FFFF")
#====================== mouse Hovering class ================================================
class HoverInfo(Menu):
 def __init__(self, parent, text, command=None):
   self._com = command
   Menu.__init__(self,parent, tearoff=0)
   if not isinstance(text, str):
      raise TypeError('Trying to initialise a Hover Menu with a non string type: ' + text.__class__.__name__)
   toktext=re.split('\n', text)
   for t in toktext:
      self.add_command(label = t)
      self._displayed=False
      self.master.bind("<Enter>",self.Display )
      self.master.bind("<Leave>",self.Remove )
 
 def __del__(self):
   self.master.unbind("<Enter>")
   self.master.unbind("<Leave>")
 
 def Display(self,event):
   if not self._displayed:
      self._displayed=True
      self.post(event.x_root, event.y_root)
   if self._com != None:
      self.master.unbind_all("<Return>")
      self.master.bind_all("<Return>", self.Click)
 
 def Remove(self, event):
  if self._displayed:
   self._displayed=False
   self.unpost()
  if self._com != None:
   self.unbind_all("<Return>")
 
 def Click(self, event):
   self._com()

#=================== Method to get the check out folder ====================================================
def get_chk_folder():

    count=0
    if ck_folder_entry.get():
        if not os.path.isdir(ck_folder_entry.get()):
            #MSS - Ticket Backlog Benchmark
            messagebox.showinfo('Error','Given path is not exists in your system ')
        else:
            print ('Given path is valid')
            print (ck_folder_entry.get())
            for i in os.listdir(ck_folder_entry.get()):
                if 'MSS - Ticket Backlog Benchmark' in i:
                    print ('Base file exists')
                    count=1
                    file_name=i
                   
            if  count==0:
                    messagebox.showinfo('Error',' Base file MSS - Ticket Backlog Benchmark not exists in the specified path')
            download_path=copy_paste.get_download_path()
            print ('Got the downbload path')
            copy_paste.move_file(download_path,ck_folder_entry.get())
            print ('Files transfer is completed')
            #time.sleep(10)
            #copy_paste.delete(ck_folder_entry.get())
    else:
        messagebox.showinfo('Error','Path Can not be Empty ')
    print ('Starting the operation')
    path=ck_folder_entry.get()
    print (file_name)
    print (path)
    final_flow.start_excel(path,file_name)
    
    #for i in os.listdir(ck_folder_entry.get()):
        #if '16th' in i:
            #print (os.path.join(ck_folder_entry.get(),i))

#======================= change the window to the center  ===============================================       
def center_window(width, height):
    # get screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # calculate position x and y coordinates
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    root.geometry('%dx%d+%d+%d' % (width, height, x, y))

#================== Method to insert the data into the data base ====================================================
def insert_yes_no():
    global entry
    if entry.get():
        result=messagebox.askquestion("Insert", "Are You Sure want to insert?")
        if result=='yes':
            creating_db.insert(entry.get())
            
            print ('Data is inserted')
            
        else:
            print ('INsert is not done')
    else:
        print ('Can not insert Null value')
        messagebox.showinfo('Error','Can not insert the NULL value')

#================== Method to delete the data into the data base ====================================================
def delete_yes_no():
    global delete_entry
    if delete_entry.get():
        result=messagebox.askquestion("Delete", "Are You Sure want to Delete?")
        if result=='yes':
             creating_db.delete(delete_entry.get())
             print ('Data is deleted')
        else:
            print ('data is not deleted')
    else:
        print ('Can not delete Null value')
        messagebox.showinfo('Error','Can not delete the NULL value')
#=========================== To destroy the addname window ==========================================
def insert_Quit():    
    global window
    window.withdraw()
#========================= To destroy the delete name window=============================================
def delete_Quit():
    global delete_window
    delete_window.withdraw()
#=========================== Mehtod to create the Add name window ==========================================
def add_name():
    global entry
    global window
    window = Toplevel(root)
    window.title('Add New name')
    name=Label(window,text='Name')
    name.grid(column=0,row=0)
    entry=Entry(window,bd=2)
    entry.grid(column=1,row=0)
    ok=Button(window,text='Ok',command=insert_yes_no)
    ok.grid(column=1,row=1)
    cancel=Button(window,text='Cancel',fg='blue',command=insert_Quit)
    cancel.grid(column=2,row=1)
#===================== Mehtod to create the Delete name window =================================================   
def delete_name():
    global delete_entry
    global delete_window
    delete_window = Toplevel(root)
    delete_window.title('Delete name')
    delete_name=Label(delete_window,text='Name')
    delete_name.grid(column=0,row=0)
    delete_entry=Entry(delete_window,bd=2)
    delete_entry.grid(column=1,row=0)
    delete_ok=Button(delete_window,text='Ok',command=delete_yes_no)
    delete_ok.grid(column=1,row=1)
    delete_cancel=Button(delete_window,text='Cancel',fg='blue',command=delete_Quit)
    delete_cancel.grid(column=2,row=1)
#========================== Metod for excel manipulation ============================================   
def excel_manipulation():
    get_chk_folder()
    #creating_db.Paetec_Closed_query()
    #creating_db.paetec_request_query()
    #creating_db.paetec_ris_query()
#====================== Method to download the files ================================================
def download_files():
    print ('Calling the method to download the files')
    open_ticket=''' ('Assigned Group*+' = "IT-OSS - M6 (ASAP/TSG)") AND ('Case Type*' = "Incident") AND (NOT('Status*' = "Resolved" OR 'Status*' = "Closed" OR 'Status*' = "Cancelled")) '''
    closed_ticket='''('Assigned Group*+' = "IT-OSS - M6 (ASAP/TSG)") AND ('Case Type*' = "Incident") AND (('Last Resolved Date' >= "10/01/2018" AND 'Last Resolved Date' < "11/01/2018"))'''
    Request_Open='''('Assigned Group*+' = "IT-OSS - M6 (ASAP/TSG)") AND ('Case Type*' = "Request") AND (NOT('Status*' = "Resolved" OR 'Status*' = "Closed" OR 'Status*' = "Cancelled"))'''
    paetec_open=creating_db.Paetec_open_query()
    #print (paetec_open)
    paetec_closed=creating_db.Paetec_Closed_query()
    paetec_request=creating_db.paetec_request_query()
    paetec_ris=creating_db.paetec_ris_query()
    query = [
        open_ticket,
        closed_ticket,
        Request_Open,
	paetec_open,
        paetec_closed,
        paetec_request,
        paetec_ris
        ]
    Backlog_report_generator_loop.download_file(query)
    

#=====================Method for place holder=================================================
def on_entry_click(event):
    if ck_folder_entry.get() == ':Specify the chekout folder':
       ck_folder_entry.delete(0, "end") # delete all the text in the entry
       ck_folder_entry.insert(0, '')
       ck_folder_entry.configure(fg='black')
#======================== Method for place holder==============================================
def on_focusin(event):
    ck_folder_entry.delete(0, "end")
    ck_folder_entry.config(fg='black')
#========================== Method for place holder============================================
def on_focusout(event):
    if ck_folder_entry.get() == '':
        ck_folder_entry.insert(0, ":Specify the chekout folder")

#==================== Labels and buttons for the main window ===============================================

center_window(500, 400)
ck_folder_label=Label(root,text='CheckOut_Folder',bg="#E0FFFF",fg='blue')
ck_folder_label.grid(column=0,row=0)

ck_folder_entry=Entry(root,bd=5,
                      width=75,fg='grey'
                      )
ck_folder_entry.grid(row=0,column=1,columnspan=4)
ck_folder_entry.insert(0,":Specify the chekout folder")
#ck_folder_entry.bind('<Key>', on_entry_click)
ck_folder_entry.bind('<FocusIn>', on_focusin)
ck_folder_entry.bind('<FocusOut>', on_focusout)


Add_name=Button(root,text='Add Name',command=add_name)
Add_name.grid(row=1,column=3)
HoverInfo(Add_name,'if u want we can add new employee name in the database ')
delete_name=Button(root,text='Delete Name',command=delete_name)
delete_name.grid(row=1,column=4,padx=5)
HoverInfo(delete_name,'if u want we can delete employee name from the database ')
go=Button(root,text='Excel Manipulation',fg='blue',command=excel_manipulation)
go.grid(row=1,column=1)
HoverInfo(go,'Run the excel Manipulation')
files_donload=Button(root,text='Download Files',fg='blue',
                     command=download_files
                     )
files_donload.grid(row=1,column=2)
HoverInfo(files_donload,'Download all the necessary reports using selenium')
root.protocol("WM_DELETE_WINDOW", creating_db.close_connection)
