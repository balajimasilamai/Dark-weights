import win32com.client as win32
import xlwings as xw
import win32api
import time
import csv

##########################
#Rename the check out file
import os
import datetime

print ("""Printing from final copy
 #########################################
  """)
path='D:\\testing-backlog-report'
month_number=(datetime.date.today()).month
day_number=(datetime.date.today()).day
year_number=(datetime.date.today()).year
print (month_number)
print (day_number)
print (year_number)
current_day_time=datetime.datetime.now()
current_day=current_day_time.strftime("%A")
print (current_day_time.strftime("%A"))
date_filed=str(month_number)+'-'+str(day_number)+'-'+str(year_number)
print (date_filed)

for i in os.listdir(path):
    if 'MSS - Ticket Backlog Benchmark_' in i:
        #os.rename(path+'\\'+i)
        print (path+'\\'+i)
        #os.rename(path+'\\'+i,path+'\\'+'MSS - Ticket Backlog Benchmark'+'_'+str(year_number)+' '+str(day_number)+' '+str(month_number)+'.xlsm')
        print ('renamed is done')
        #file_name='MSS - Ticket Backlog Benchmark'+'_'+str(year_number)+' '+str(day_number)+' '+str(month_number)+'.xlsm'
        file_name=i
        print (path+'\\'+file_name)

#########################
def start_excel(path,file_name):
    #Get the checkout folder path
    #path='D:\\testing-backlog-report'
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = True
    wb = excel.Workbooks.Open(path+'\\'+file_name)
    ws = wb.Worksheets("Highlevel Bench Mark")
    lastCol = ws.UsedRange.Columns.Count
    lastRow = ws.UsedRange.Rows.Count
    print (lastCol)
    print (lastRow)
    print (ws.Cells(2, lastCol).Value)
    print (ws.Cells(3, lastCol).Value)
    # create a new column in Highlevel Bench Mark sheet
    excel.Application.Run("Macro_copy")
    time.sleep(2)
    lastCol_new = ws.UsedRange.Columns.Count
    print (lastCol_new)
    print (ws.Cells(2, lastCol_new).Value)
    print (ws.Cells(3, lastCol_new).Value)
    #updating the date and day cells in Highlevel Bench Mark sheet
    ws.Cells(2, lastCol_new).Value=current_day
    ws.Cells(3, lastCol_new).Value=date_filed

    '''
    # Copy the values for the PAETEC Request sheet
    ws = wb.Worksheets("PAETEC Request")
    #This loop for PAETEC Request sheet
    row=1
    col=1
    with open(path+'paetec_request_22thOct18.csv','r') as csv_file:
        c=csv.reader(csv_file)
        for i,val in enumerate(c):
            if i != 0:
                for num in range(0,len(val)):
                    #print (num)        
                    #print (i,val)
                    ws.Cells(row,col).Value=val[num]
                    col=col+1
            row=row+1
            col=1 
    #This loop for PAETEC RIS sheet
    ws = wb.Worksheets("PAETEC RIS")
    row=1
    col=1
    with open(path+'paetec_ris_22thOct18.csv','r') as csv_file:
        c=csv.reader(csv_file)
        for i,val in enumerate(c):
            if i != 0:
                for num in range(0,len(val)):
                    #print (num)        
                    #print (i,val)
                    ws.Cells(row,col).Value=val[num]
                    col=col+1
            row=row+1
            col=1
    '''
    # Go to the calc sheet and update the path
    ws = wb.Worksheets("Calc")
    ws.Cells(5,2).Value = path
    #Now we have done all the neccesarry activity  so call the macro to populate the values in respective fields'
    excel.Application.Run("Macro6")
    time.sleep(1)
    excel.Application.Run("macro_cal_save")
    ###############################
    # verication of the metrics count with high level bench mark
    tc_received=ws.Cells(19,2).Value
    tc_resolved_same_day=ws.Cells(21,2).Value
    pae_tc_recieved=ws.Cells(27,2).Value
    pae_tc_resolved_same_day=ws.Cells(29,2).Value
    # To get the values from Metrics worksheet
    wc = wb.Worksheets("Metrics")
    metrics=wc.Cells(23,6).Value
    # To get the values from Metrics - Paetec worksheet
    wc = wb.Worksheets("Metrics - Paetec")
    metrics_pae=wc.Cells(13,6).Value#ws.Cells(23,lastCol+1).Value
    #lastCol = ws.UsedRange.Columns.Count
    #lastRow = ws.UsedRange.Rows.Count
    ws = wb.Worksheets("Highlevel Bench Mark")

    #To get the values for hwin count
    a=ws.Cells(16,lastCol_new-1).Value
    b=ws.Cells(4,lastCol_new-1).Value
    c=ws.Cells(6,lastCol_new-1).Value
    d=ws.Cells(5,lastCol_new-1).Value
    print ('ws.Cells(16,lastCol_new-1) ',a)
    print ('ws.Cells(4,lastCol_new-1) ',b)
    print ('ws.Cells(6,lastCol_new-1) ',c)
    print ('ws.Cells(5,lastCol_new-1)',d)
    #BKR26+BKS22-BKS23-BKS24
    #Get the values for paetec_metrics
    pa=ws.Cells(26,lastCol_new).Value
    pb=ws.Cells(22,lastCol_new).Value
    pc=ws.Cells(23,lastCol_new).Value
    pd=ws.Cells(24,lastCol_new).Value

    print ('ws.Cells(28,lastCol_new)  ',pa)
    print ('ws.Cells(22,lastCol_new) ',pb)
    print ('ws.Cells(23,lastCol_new) ',pc)
    print ('ws.Cells(24,lastCol_new) ',pd)


    actual_hwin_val=ws.Cells(16,lastCol_new).Value
    print ('actual_hwin_val ',actual_hwin_val)
    print ('metrics ',metrics)
    actual_hpae_val=ws.Cells(26,lastCol_new).Value
    print ('actual_hpae_val ',actual_hpae_val)
    print ('metrics_pae ', metrics_pae)
    #print (pa)
    #print (pb)
    #print (pc)
    #print (pd)
    #Checking the count for hwin
    if actual_hwin_val != metrics:
            if a+b-(c+d) < metrics:
                    diff=metrics - actual_hwin_val 
                    correct_c_value=c-diff
                    print (' correct_c_value ',correct_c_value)
            else:
                    diff=actual_hwin_val - metrics
                    correct_c_value=c+diff
                    print (' correct_c_value ',correct_c_value)
    else:
            correct_c_value=c
            print (' correct_c_value ',correct_c_value)

    #Updating the correct_a_value
    try:
            ws.Cells(6,lastCol_new-1).Value=correct_c_value
    except Exception as e:
            print (e)

    #Checking the count for paetec
    if actual_hpae_val != metrics_pae:
            if actual_hpae_val < metrics_pae:
                    diff=metrics_pae - actual_hpae_val
                    correct_pd_value=pd-diff
                    print ('correct_pd_value ',correct_pd_value)
            else:
                    diff=actual_hpae_val - metrics_pae
                    correct_pd_value=pd+diff
                    print ('correct_pd_value ',correct_pd_value)
    else:
            correct_pd_value=pd
            print ('correct_pd_value ',correct_pd_value)
    #Updating the correct_pa_value
    try:
            ws.Cells(24,lastCol_new).Value=correct_pd_value
    except Exception as e:
            print (e)
    ################
    #To close the opened the file
    excel.Application.Run("macro_save")
    wb.Close()
    excel.Quit()
#start_excel(path,file_name)

