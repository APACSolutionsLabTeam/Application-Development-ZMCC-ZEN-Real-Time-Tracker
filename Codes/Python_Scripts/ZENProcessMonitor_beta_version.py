# -*- coding: utf-8 -*-
"""
Created on Fri Mar 26 14:13:24 2021

@author: ZSPANIYA
"""
import psutil
import time
import csv
import os
import pywintypes
from win10toast import ToastNotifier
from datetime import datetime, timedelta
import pickle
import pandas as pd

import os
if not os.path.exists('C:\ZENProcessMonitorLogs'):
    os.makedirs('C:\ZENProcessMonitorLogs')
    
ZENBlue_pickle_filename = 'C:\ZENProcessMonitorLogs\ZENBlue_picklefile.pickle'
ZENCore_pickle_filename = 'C:\ZENProcessMonitorLogs\ZENCore_picklefile.pickle'
SmartSEM_pickle_filename = 'C:\ZENProcessMonitorLogs\SmartSEM_picklefile.pickle'

ZEN_Blue = 'ZEN.exe' 
ZEN_Core = 'ZENCore.exe' 
SmartSEM = 'SmartSEM.exe'
ls = []
csvFilename_ZENBlue = 'C:\ZENProcessMonitorLogs\ZEN_Blue_Tracker.csv'
csvFilename_ZENCore = 'C:\ZENProcessMonitorLogs\ZEN_Core_Tracker.csv'
csvFilename_SmartSEM = 'C:\ZENProcessMonitorLogs\SmartSEM_Tracker.csv'
process_time= {}
timestamp = {}
SSid_Blue = 0
SSid_Core = 0
SSid_SmartSEM = 0
stop_time = 'Still Running'
toast = ToastNotifier()
toast.show_toast("ZEN Software Real-Time Tracker", "The ZEN Process Monitor has started", 'ZEISSicon_48X48.ICO', duration = 30)
run_Blue_count = 0
run_Core_count = 0
run_SmartSEM_count = 0

#os.chdir(r'C:\Users\ZSPANIYA\Desktop\ZENMonitor')
while True:
   
    with open (csvFilename_ZENBlue, 'a+') as csvfile_ZENBlue:
        headers = ["Name", "Start Time", "Stop Time","Active Time in DayHourMinSec", "Last Check Time", "Username", "Session ID", "Total Activity Time"]
        writer = csv.DictWriter(csvfile_ZENBlue, delimiter=',', lineterminator='\n',fieldnames=headers)
        
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        if csvfile_ZENBlue.tell() == 0:
            writer.writeheader()  # file doesn't exist yet, write a header
            
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
                    SSid_Blue = SSid_Blue + 1    
                    start_time = datetime.fromtimestamp(p.create_time()).strftime("%d/%m/%Y %H:%M:%S")
                    with open(ZENBlue_pickle_filename, 'wb') as f:
                        pickle.dump(p.pid, f)
                    
                    if (run_Blue_count!=1):
                        writer.writerow({'Name': 'ZEN Blue', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_Blue)})                    
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
                        writer.writerow({'Name': 'ZEN Blue', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_Blue)})
        process_status_ZENBlue = [ p.status() for p in psutil.process_iter() if p.name() == ZEN_Blue ]    
        #print(run_Blue_count)
        if process_status_ZENBlue == []:
            if csvfile_ZENBlue.tell() != 0 and run_Blue_count>0:
                #print('CSV file present')
                df = pd.read_csv(csvFilename_ZENBlue)
                # updating the column value/data
                df.loc[df.shape[0]-1, 'Stop Time'] = df.loc[df.shape[0]-1, 'Last Check Time']
                df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                #df.loc[0,'Total Activity Time'] = df['Active Time in DayHourMinSec'].sum()
                df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                df.to_csv(csvFilename_ZENBlue, index=False)
                # print( df['Active Time in DayHourMinSe'].cumsum())
               
                
                
       
                
                
             
    with open (csvFilename_ZENCore, 'a+') as csvfile_ZENCore:
        headers = ["Name", "Start Time", "Stop Time", "Active Time in DayHourMinSec", "Last Check Time", "Username", "Session ID", "Total Activity Time"]
        writer = csv.DictWriter(csvfile_ZENCore, delimiter=',', lineterminator='\n',fieldnames=headers)
        
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        if csvfile_ZENCore.tell() == 0:
            writer.writeheader()  # file doesn't exist yet, write a header
            
    
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
                #print('pid', p.pid)
                #print('pid_old', pid_old)
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
                        writer.writerow({'Name': 'ZEN Core', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_Core)})                    
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
                        writer.writerow({'Name': 'ZEN Core', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_Core)})

        process_status_ZENCore = [ p.status() for p in psutil.process_iter() if p.name() == ZEN_Core ]    
        #print(process_status_ZENCore)
        #print(run_Core_count)
        if process_status_ZENCore == []:
            if csvfile_ZENCore.tell() != 0 and run_Core_count>0:
                #print('CSV file present')
                df = pd.read_csv(csvFilename_ZENCore)
                # updating the column value/data
                df.loc[df.shape[0]-1, 'Stop Time'] = df.loc[df.shape[0]-1, 'Last Check Time']
                df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                df.to_csv(csvFilename_ZENCore, index=False)                               
    
    with open (csvFilename_SmartSEM, 'a+') as csvfile_SmartSEM:
        headers = ["Name", "Start Time", "Stop Time","Active Time in DayHourMinSec", "Last Check Time", "Username", "Session ID", "Total Activity Time"]
        writer = csv.DictWriter(csvfile_SmartSEM, delimiter=',', lineterminator='\n',fieldnames=headers)
        
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M")
        if csvfile_SmartSEM.tell() == 0:
            writer.writeheader()  # file doesn't exist yet, write a header
            
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
                        writer.writerow({'Name': 'SmartSEM', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_SmartSEM)})                    
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
                        writer.writerow({'Name': 'SmartSEM', 'Start Time': start_time, 'Stop Time': stop_time, 'Active Time in DayHourMinSec':proc_time, 'Last Check Time': dt_string, 'Username': process_user, 'Session ID': 'S000'+str(SSid_SmartSEM)})
        process_status_SmartSEM = [ p.status() for p in psutil.process_iter() if p.name() == SmartSEM ]    
        #print(run_Blue_count)
        if process_status_SmartSEM == []:
            if csvfile_SmartSEM.tell() != 0 and run_SmartSEM_count>0:
                #print('CSV file present')
                df = pd.read_csv(csvFilename_SmartSEM)
                # updating the column value/data
                df.loc[df.shape[0]-1, 'Stop Time'] = df.loc[df.shape[0]-1, 'Last Check Time']
                df['Active Time in DayHourMinSec'] = pd.to_timedelta(df['Active Time in DayHourMinSec'])
                df.loc[0,'Total Activity Time']  = df['Active Time in DayHourMinSec'].sum()
                df.to_csv(csvFilename_SmartSEM, index=False)
                
                
    
    time.sleep(180)
#toast.show_toast("ZEN Software Real-Time Tracker", "The ZEN Process Monitor has stopped", icon_path='', duration = 30)
