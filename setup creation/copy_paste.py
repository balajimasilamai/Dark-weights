import win32api
#import datetime
import os
import shutil
import re

#check_out_folder='D:\\testing-backlog-report\\'
#current_time=datetime.date.today().strftime("%B %dth, %Y")
#month=datetime.date.today().strftime("%B")
#day=datetime.date.today().strftime("%dth")
#year=datetime.date.today().strftime("%Y")
# Get the downloaded file path
def get_download_path():
        name=win32api.GetUserNameEx(win32api.NameSamCompatible)
        rename=re.sub(r'\W', " ", name)
        #print (rename)
        pos=rename.find(' ')
        #print (pos)
        path2='Downloads'
        path1='Users'
        #print (name)
        #print (name[8:])        
        user_name=name[pos+1:]
        #print ('user name :'+name)
        full_path='C:'+"\\"+path1+"\\"+user_name+"\\"+path2
        print (full_path)
        return full_path
#To move the file from download path to the checkout folder
def move_file(get_download_path,to_path,dt_time):
    #file_name=day[0:]+month[0:3]+year[2:4]
    #print (file_name)
    #file_name1='10thOct'
    #if not os.path.isdir(to_path):
        #os.mkdir(to_path)
        #print ('dir created')
    #else:
        #print ('Already exists ')
    for i in os.listdir(get_download_path):
        if dt_time in i:
            print (i)
            print (to_path+i)
            shutil.copy(os.path.join(get_download_path,i), os.path.join(to_path,i))
def delete(delete_path):
        #delete_path='D:\\testing-backlog-report\\'
        for file in os.listdir(delete_path):
                if '.csv' in file:
                        os.remove(os.path.join(delete_path,file))
                        print (file +' is deleted')


if '_main_'=='_main_':
    #get_path=get_download_path()
    #print (get_path)
    #move_file(get_path,check_out_folder)
    pass
