# -*- coding: utf-8 -*-
"""
Created on Tue Jan 19 14:22:22 2021

@author: ZSPANIYA
"""

import psutil
import time
import csv
import os
import pywintypes
from win10toast import ToastNotifier
from datetime import datetime




ZEN_Blue = 'ZEN.exe' 
ZEN_Core = 'ZENCore.exe' 
ls = []
csvFilename = 'C:\Temp\ZEN_Monitor.csv'

toast = ToastNotifier()
toast.show_toast("ZEN Software Real-Time Tracker", "The ZEN Process Monitor has started", duration = 30)

#os.chdir(r'C:\Users\ZSPANIYA\Desktop\ZENMonitor')
while True:
    with open (csvFilename, 'a+') as csvfile:
        headers = ["Process ID", "Name", "Status", "Started Time", "Check Time"]
        writer = csv.DictWriter(csvfile, delimiter=',', lineterminator='\n',fieldnames=headers)
        now = datetime.now()
        #current_time = now.strftime("%H:%M:%S")
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        print("date and time =", dt_string)
        if csvfile.tell() == 0:
            writer.writeheader()  # file doesn't exist yet, write a header
    
        for p in psutil.process_iter(['name']):
            
            if p.info['name'] == ZEN_Blue:
                ls.append(p)
                # print(p.create_time())
                # print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(p.create_time())))
                # writer.writerow({'Process ID': p.pid, 'Name': p.info['name'], 'Status': p.status(), 'Started Time': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(p.create_time()))})
                writer.writerow({'Process ID': p.pid, 'Name': 'ZEN Blue', 'Status': p.status(), 'Started Time': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(p.create_time())), 'Check Time': dt_string})
                
            if p.info['name'] == ZEN_Core:
                ls.append(p)  
                writer.writerow({'Process ID': p.pid, 'Name': 'ZEN Core', 'Status': p.status(), 'Started Time': time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(p.create_time())), 'Check Time': dt_string})
            if(ls==[]):   
                continue
    
    time.sleep(300)
toast.show_toast("ZEN Software Real-Time Tracker", "The ZEN Process Monitor has stopped", duration = 30)
dt_string2 = now.strftime("%d/%m/%Y %H:%M:%S")