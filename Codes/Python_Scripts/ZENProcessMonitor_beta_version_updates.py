# -*- coding: utf-8 -*-
"""
Created on Fri Mar 26 14:13:24 2021

@author: ZSPANIYA
"""
## ************************** Importing all required libraries  *****************************
import psutil
import time
import csv
import os
from win10toast import ToastNotifier
from datetime import datetime, timedelta
import pickle
import pandas as pd
import re
import geocoder
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import tkinter as tk
from tkinter import simpledialog
import socket

## *********************** Simple GUI for fetching system name from the user (only once) *****************
root = tk.Tk() # Create an instance of tkinter

system_name = simpledialog.askstring(title = "System Name",
                                    prompt = "Please enter the system name for which the log is generated:")

print(system_name)
root.destroy()

## ************************* Check internet connectivity *************************************
## getting the hostname by socket.gethostname() method
hostname = socket.gethostname()
print(hostname)
## getting the IP address using socket.gethostbyname() method
ip_address = socket.gethostbyname(hostname)


## *********************** If internet connection available => allow copy to google drive and get geolocation *************
if ip_address=="127.0.0.1":
    print("No internet, your localhost is "+ ip_address)
else:
    print("Connected, with the IP address: "+ ip_address )
    # printing the hostname and ip_address
    print(f"Hostname: {hostname}")
    print(f"IP Address: {ip_address}")

if ip_address != "127.0.0.1": 
    gauth = GoogleAuth()
    # Try to load saved client credentials
    gauth.LoadCredentialsFile("Trackercreds.txt")
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()
    # Save the current credentials to a file
    gauth.SaveCredentialsFile("Trackercreds.txt")
    
    drive = GoogleDrive(gauth)

## **************************************** Get geolocation **********************************************        
    device_loc = geocoder.ip('me') # actually need to provide IP address - works for public networks
    device_city = device_loc.city
    device_country = device_loc.country
    device_lat_long = device_loc.latlng

## *********************** Define log destination directory (with system name) and filenames ***************************************
LogDirectory = 'C:\ZEN_SmartSEM_MonitorLogs_'+system_name
if not os.path.exists(LogDirectory):
    os.makedirs(LogDirectory)
    
ZENBlue_pickle_filename = LogDirectory+'\ZENBlue_picklefile.pickle'
ZENCore_pickle_filename = LogDirectory+'\ZENCore_picklefile.pickle'
SmartSEM_pickle_filename = LogDirectory+'\SmartSEM_picklefile.pickle'

## ********************* Define exe of applications to track and other global variables ***********************************
ZEN_Blue = 'ZEN.exe' 
ZEN_Core = 'ZENCore.exe' 
SmartSEM = 'SmartSEM.exe'
ls = []
csvFilename_ZENBlue = LogDirectory+'\ZEN_Blue_Tracker.csv'
csvFilename_ZENCore = LogDirectory+'\ZEN_Core_Tracker.csv'
csvFilename_SmartSEM = LogDirectory+'\SmartSEM_Tracker.csv'
process_time= {}
timestamp = {}
SSid_Blue = 0
SSid_Core = 0
SSid_SmartSEM = 0
stop_time = 'Still Running'
toast = ToastNotifier()
toast.show_toast("ZEN SmartSEM Software Real-Time Tracker", "The ZEN SmartSEM Monitor has started", "C:\ZEN_SmartSEM_MonitorExecutables\ZEISSicon_48X48.ICO", duration = 30)
run_Blue_count = 0
run_Core_count = 0
run_SmartSEM_count = 0
copy_once_month = 0
sent_file_flag = False
sent_date_time = time.time()
##
## ******************* Running an infinite loop - continously read the windows process information ************************
while True:
    ## Tracking ZEN Blue
    with open (csvFilename_ZENBlue, 'a+') as csvfile_ZENBlue:
        headers = ["Name", "Start Time", "Stop Time","Active Time in DayHourMinSec", "Last Check Time", "Username", "Session ID", "Geolocation(CITY_COUNTRY)", "Geolocation(LAT_LONG)", "Total Activity Time"]
        writer = csv.DictWriter(csvfile_ZENBlue, delimiter=',', lineterminator='\n',fieldnames=headers)
        
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        if csvfile_ZENBlue.tell() == 0:
            writer.writeheader()  # file doesn't exist yet, write a header
         
        process_status_ZENBlue = [ p.status() for p in psutil.process_iter() if p.name() == ZEN_Blue ]    
       
        if csvfile_ZENBlue.tell() != 0:
            df = pd.read_csv(csvFilename_ZENBlue)
            
            if (df.shape[0] !=0):
                    # updating the column value
                    ssid = df['Session ID'].iloc[-1]
                    SSid_Blue  = (int(' '.join(map(str, re.findall(r'\d+', ssid)))))
                    #print(SSid_Blue)
                   
                    if (process_status_ZENBlue == [] and df['Stop Time'].iloc[-1] == 'Still Running'):
                        df.loc[df.shape[0]-1, 'Stop Time'] = df.loc[df.shape[0]-1, 'Last Check Time']
                    df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                    #df.loc[0,'Total Activity Time'] = df['Active Time in DayHourMinSec'].sum()
                    df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                    df.to_csv(csvFilename_ZENBlue, index=False)
                    # print( df['Active Time in DayHourMinSe'].cumsum())    
            
        for p in psutil.process_iter(['name']):   
          
           if p.info['name'] == ZEN_Blue:
              #print(p.status())            
              if  p.status() == 'running':
                # 
                # print(process_status)
                run_Blue_count = run_Blue_count + 1
                #run_Blue_count_temp = run_Blue_count
                ls.append(p)
                timestamp[p] = int(p.create_time())
                process_time[p] = int(time.time())-timestamp[p]
                process_user = p.username()
                try:
                      pid_old = pickle.load(open(ZENBlue_pickle_filename, "rb")) 
                except (OSError, IOError):
                        pid_old = 0
                        pickle.dump(pid_old, open(ZENBlue_pickle_filename, "wb")) 
                proc_time = "{:0>8}".format(str(timedelta(seconds=process_time[p])))
               #print(proc_time)
                if p not in process_time.keys():
                    process_time[p] = 0
                process_time[p] = process_time[p]+int(time.time())-timestamp[p]
                #print('pid', p.pid)
                #print('pid_old', pid_old)
                if (p.pid - pid_old == 0):
                   with open(csvFilename_ZENBlue, "r") as file:
                        reader = csv.reader(file, delimiter=',', lineterminator='\n')
                        header = next(reader)
                        data = list(reader)
                        rowValues = []
                        for row in data:
                            colValues = []
                            for col in row:
                                colValues.append(col)
                            rowValues.append(colValues)
                            #print(rowValues)
                        colValues[2] = stop_time  
                        colValues[3] = proc_time
                        colValues[4] = dt_string 
                                               
                   with open(csvFilename_ZENBlue, "w") as outFile:
                        writer = csv.writer(outFile, delimiter=',', lineterminator='\n')
                        writer.writerow(header)
                        
                        for col in rowValues:
                            writer.writerow(col)
                        
                else:
                    #print(SSid_Blue)
                    SSid_Blue = SSid_Blue + 1    
                    start_time = datetime.fromtimestamp(p.create_time()).strftime("%d/%m/%Y %H:%M:%S")
                    with open(ZENBlue_pickle_filename, 'wb') as f:
                        pickle.dump(p.pid, f)
                    
                    if (run_Blue_count!=1):
                        writer.writerow({'Name': 'ZEN Blue', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_Blue), 'Geolocation(CITY_COUNTRY)':device_city+' ,'+ device_country, 'Geolocation(LAT_LONG)': str(device_lat_long)[1:-1]})                    
                        with open(csvFilename_ZENBlue, "r") as file2:
                            reader = csv.reader(file2, delimiter=',', lineterminator='\n')
                            header = next(reader)
                            data = list(reader)
                            #print(data)
                            rowValues = []
                            for row in data:
                                colValues = []
                                for col in row:
                                    colValues.append(col)
                                rowValues.append(colValues)
                            colValues[2] = colValues[4] 
                                                
                        with open(csvFilename_ZENBlue, "w") as outFile2:
                            writer = csv.writer(outFile2, delimiter=',', lineterminator='\n')
                            writer.writerow(header)
                            
                            for col in rowValues:
                                writer.writerow(col)
                        
                    else:
                        writer.writerow({'Name': 'ZEN Blue', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_Blue), 'Geolocation(CITY_COUNTRY)':device_city+' ,'+ device_country, 'Geolocation(LAT_LONG)': str(device_lat_long)[1:-1]})
        
        if csvfile_ZENBlue.tell() != 0:
            df = pd.read_csv(csvFilename_ZENBlue)
            if (df.shape[0] !=0):
                df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                #df.loc[0,'Total Activity Time'] = df['Active Time in DayHourMinSec'].sum()
                df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                df.to_csv(csvFilename_ZENBlue, index=False)   
            
              
    ## Tracking ZEN Core  
    with open (csvFilename_ZENCore, 'a+') as csvfile_ZENCore:
        headers = ["Name", "Start Time", "Stop Time", "Active Time in DayHourMinSec", "Last Check Time", "Username", "Session ID", "Geolocation(CITY_COUNTRY)", "Geolocation(LAT_LONG)", "Total Activity Time"]
        writer = csv.DictWriter(csvfile_ZENCore, delimiter=',', lineterminator='\n',fieldnames=headers)
        
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        if csvfile_ZENCore.tell() == 0:
            writer.writeheader()  # file doesn't exist yet, write a header
        
        process_status_ZENCore = [ p.status() for p in psutil.process_iter() if p.name() == ZEN_Core ]    
        if csvfile_ZENCore.tell() != 0:
            df = pd.read_csv(csvFilename_ZENCore)
            
            if (df.shape[0] !=0):
                    # updating the column value/
                    ssid = df['Session ID'].iloc[-1]
                    SSid_Core  = (int(' '.join(map(str, re.findall(r'\d+', ssid)))))
                    #print(SSid_Blue)
                   
                    if (process_status_ZENCore == [] and df['Stop Time'].iloc[-1] == 'Still Running'):
                        df.loc[df.shape[0]-1, 'Stop Time'] = df.loc[df.shape[0]-1, 'Last Check Time']
                    df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                    #df.loc[0,'Total Activity Time'] = df['Active Time in DayHourMinSec'].sum()
                    df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                    df.to_csv(csvFilename_ZENCore, index=False)
                    # print( df['Active Time in DayHourMinSe'].cumsum())
                    
        for p in psutil.process_iter(['name']):
            
            if p.info['name'] == ZEN_Core:
                run_Core_count = run_Core_count + 1
                ls.append(p)
                timestamp[p] = int(p.create_time())
                process_time[p] = int(time.time())-timestamp[p]
                process_user = p.username()
                
                try:
                      pid_old = pickle.load(open(ZENCore_pickle_filename, "rb")) 
                except (OSError, IOError):
                        pid_old = 0
                        pickle.dump(pid_old, open(ZENCore_pickle_filename, "wb"))
                        
                proc_time = "{:0>8}".format(str(timedelta(seconds=process_time[p])))
                
                if p not in process_time.keys():
                    process_time[p] = 0
                process_time[p] = process_time[p]+int(time.time())-timestamp[p]
                if (p.pid - pid_old == 0):
                    with open(csvFilename_ZENCore, "r") as file:
                        reader = csv.reader(file, delimiter=',', lineterminator='\n')
                        header = next(reader)
                        data = list(reader)
                        rowValues = []
                        for row in data:
                            colValues = []
                            for col in row:
                                colValues.append(col)
                            rowValues.append(colValues)
                        colValues[3] = proc_time
                        colValues[4] = dt_string    
                        
                    with open(csvFilename_ZENCore, "w") as outFile:
                        writer = csv.writer(outFile, delimiter=',', lineterminator='\n')
                        writer.writerow(header)
                        
                        for col in rowValues:
                            writer.writerow(col)
                        
                else:
                    SSid_Core = SSid_Core + 1    
                    start_time = datetime.fromtimestamp(p.create_time()).strftime("%d/%m/%Y %H:%M:%S")
                    with open(ZENCore_pickle_filename, 'wb') as f:
                        pickle.dump(p.pid, f)
                    
                    if (run_Core_count!=1):
                        writer.writerow({'Name': 'ZEN Core', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_Core), 'Geolocation(CITY_COUNTRY)':device_city+' ,'+ device_country, 'Geolocation(LAT_LONG)': str(device_lat_long)[1:-1]})                    
                        with open(csvFilename_ZENCore, "r") as file2:
                            reader = csv.reader(file2, delimiter=',', lineterminator='\n')
                            header = next(reader)
                            data = list(reader)
                            #print(data)
                            rowValues = []
                            for row in data:
                                colValues = []
                                for col in row:
                                    colValues.append(col)
                                rowValues.append(colValues)
                            colValues[2] = colValues[4] 
                                                
                        with open(csvFilename_ZENCore, "w") as outFile2:
                            writer = csv.writer(outFile2, delimiter=',', lineterminator='\n')
                            writer.writerow(header)
                            
                            for col in rowValues:
                                writer.writerow(col)
                        
                    else:
                        writer.writerow({'Name': 'ZEN Core', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_Core), 'Geolocation(CITY_COUNTRY)':device_city+' ,'+ device_country, 'Geolocation(LAT_LONG)': str(device_lat_long)[1:-1]})

        if csvfile_ZENCore.tell() != 0:
            df = pd.read_csv(csvFilename_ZENCore)
            if (df.shape[0] !=0):
                df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                #df.loc[0,'Total Activity Time'] = df['Active Time in DayHourMinSec'].sum()
                df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                df.to_csv(csvFilename_ZENCore, index=False)   
            
                                 
    ## Tracking SmartSEM 
    with open (csvFilename_SmartSEM, 'a+') as csvfile_SmartSEM:
        headers = ["Name", "Start Time", "Stop Time","Active Time in DayHourMinSec", "Last Check Time", "Username", "Session ID", "Geolocation(CITY_COUNTRY)", "Geolocation(LAT_LONG)", "Total Activity Time"]
        writer = csv.DictWriter(csvfile_SmartSEM, delimiter=',', lineterminator='\n',fieldnames=headers)
        
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M")
        if csvfile_SmartSEM.tell() == 0:
            writer.writeheader()  # file doesn't exist yet, write a header
            
        process_status_SmartSEM = [ p.status() for p in psutil.process_iter() if p.name() == SmartSEM ]    
        if csvfile_SmartSEM.tell() != 0:
            df = pd.read_csv(csvFilename_SmartSEM)
            
            if (df.shape[0] !=0):
                    # updating the column value/
                    ssid = df['Session ID'].iloc[-1]
                    SSid_SmartSEM  = (int(' '.join(map(str, re.findall(r'\d+', ssid)))))
                    #print(SSid_Blue)
                   
                    if (process_status_SmartSEM == [] and df['Stop Time'].iloc[-1] == 'Still Running'):
                        df.loc[df.shape[0]-1, 'Stop Time'] = df.loc[df.shape[0]-1, 'Last Check Time']
                    df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                    #df.loc[0,'Total Activity Time'] = df['Active Time in DayHourMinSec'].sum()
                    df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                    df.to_csv(csvFilename_SmartSEM, index=False)   
            
        for p in psutil.process_iter(['name']):   
          
           if p.info['name'] == SmartSEM:
              #print(p.status())            
              if  p.status() == 'running':
                # 
                # print(process_status)
                run_SmartSEM_count = run_SmartSEM_count + 1
                #run_Blue_count_temp = run_Blue_count
                ls.append(p)
                timestamp[p] = int(p.create_time())
                process_time[p] = int(time.time())-timestamp[p]
                process_user = p.username()
                try:
                      pid_old = pickle.load(open(SmartSEM_pickle_filename, "rb")) 
                except (OSError, IOError):
                        pid_old = 0
                        pickle.dump(pid_old, open(SmartSEM_pickle_filename, "wb")) 
                proc_time = "{:0>8}".format(str(timedelta(seconds=process_time[p])))
               #print(proc_time)
                if p not in process_time.keys():
                    process_time[p] = 0
                process_time[p] = process_time[p]+int(time.time())-timestamp[p]
                #print('pid', p.pid)
                #print('pid_old', pid_old)
                if (p.pid - pid_old == 0):
                   with open(csvFilename_SmartSEM, "r") as file:
                        reader = csv.reader(file, delimiter=',', lineterminator='\n')
                        header = next(reader)
                        data = list(reader)
                        rowValues = []
                        for row in data:
                            colValues = []
                            for col in row:
                                colValues.append(col)
                            rowValues.append(colValues)
                            #print(rowValues)
                        colValues[2] = stop_time  
                        colValues[3] = proc_time
                        colValues[4] = dt_string 
                                               
                   with open(csvFilename_SmartSEM, "w") as outFile:
                        writer = csv.writer(outFile, delimiter=',', lineterminator='\n')
                        writer.writerow(header)
                        
                        for col in rowValues:
                            writer.writerow(col)
                        
                else:
                    SSid_SmartSEM = SSid_SmartSEM + 1    
                    start_time = datetime.fromtimestamp(p.create_time()).strftime("%d/%m/%Y %H:%M:%S")
                    with open(SmartSEM_pickle_filename, 'wb') as f:
                        pickle.dump(p.pid, f)
                    
                    if (run_SmartSEM_count!=1):
                        writer.writerow({'Name': 'SmartSEM', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_SmartSEM), 'Geolocation(CITY_COUNTRY)':device_city+' ,'+ device_country, 'Geolocation(LAT_LONG)': str(device_lat_long)[1:-1]})                    
                        with open(csvFilename_SmartSEM, "r") as file2:
                            reader = csv.reader(file2, delimiter=',', lineterminator='\n')
                            header = next(reader)
                            data = list(reader)
                            #print(data)
                            rowValues = []
                            for row in data:
                                colValues = []
                                for col in row:
                                    colValues.append(col)
                                rowValues.append(colValues)
                            colValues[2] = colValues[4] 
                                                
                        with open(csvFilename_SmartSEM, "w") as outFile2:
                            writer = csv.writer(outFile2, delimiter=',', lineterminator='\n')
                            writer.writerow(header)
                            
                            for col in rowValues:
                                writer.writerow(col)
                        
                    else:
                        writer.writerow({'Name': 'SmartSEM', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_SmartSEM), 'Geolocation(CITY_COUNTRY)':device_city+' ,'+ device_country, 'Geolocation(LAT_LONG)': str(device_lat_long)[1:-1]})
       
        if csvfile_SmartSEM.tell() != 0:
            df = pd.read_csv(csvFilename_SmartSEM)
            if (df.shape[0] !=0):
                df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                #df.loc[0,'Total Activity Time'] = df['Active Time in DayHourMinSec'].sum()
                df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                df.to_csv(csvFilename_SmartSEM, index=False)   
    
    hostname = socket.gethostname()
    ## getting the IP address using socket.gethostbyname() method
    ip_address = socket.gethostbyname(hostname)
    #twoweekslater = sent_date_time + datetime.timedelta(days=1)
    #if datetime.now() >= twoweekslater and sent_file_flag is False:
    #if time.time()-sent_date_time>=1209600:# 2 weeks
    #if time.time()-sent_date_time>=3600:# 1 hour
    if ip_address != "127.0.0.1": 
            
            ZENBlueLogfile = drive.CreateFile()
            ZENBlueLogfile.SetContentFile(LogDirectory+'\ZEN_Blue_Tracker.csv')
            ZENBlueLogfile.Upload()
            
            ZENCoreLogfile = drive.CreateFile()
            ZENCoreLogfile.SetContentFile(LogDirectory+'\ZEN_Core_Tracker.csv')
            ZENCoreLogfile.Upload()
            
            SmartSEMLogfile = drive.CreateFile()
            SmartSEMLogfile.SetContentFile(LogDirectory+'\SmartSEM_Tracker.csv')
            SmartSEMLogfile.Upload() 
            
            sent_file_flag = True
            sent_date_time = time.time()
    else:
             sent_file_flag = False
                  
    time.sleep(60)
#toast.show_toast("ZEN Software Real-Time Tracker", "The ZEN Process Monitor has stopped", icon_path='', duration = 30)
