# -*- coding: utf-8 -*-
"""
Created on Wed May 20 13:22:29 2015

@author: Don
"""

#! /usr/bin/python

import datetime
import time
import csv
import smtplib
import os, sys
import itertools

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

   
def processData():
    
    #partialfilename = "C:\myFolder\StockChange"
    
    ordersDataList = []
    
    ordersDataList.append(getListingCount())   #HourlyListingCount
      
    emailSubject = constructEmail("S",ordersDataList)
    emailBody    = constructEmail("B",ordersDataList)
    
    sendEmail(emailSubject,emailBody)
    
    return

def cleanData():
    
    partialfilename = "C:\myFolder\StockChange"
    lstngTimeStampPrev = " "
        
    fib1  = open(partialfilename+".csv")    
    fi_b1 = csv.reader(fib1)
    
    fob1  = open(partialfilename+"Clean.csv",'w')
    fo_b1 = csv.writer(fob1,lineterminator='\n')

    for row1 in fi_b1:
        
        if row1[3] == lstngTimeStampPrev:
            x = 0
        else:
            lstngTimeStampPrev = row1[3]
            fo_b1.writerow(row1)

    fib1.close()
    fob1.close()
        
    return 
    
def getListingCount():
    
    partialfilename = "C:\myFolder\StockChange"
    columnSKU       = 0
    columnStkNow    = 1
    columnLvlChg    = 2    
    columnDate      = 3
    columnLstr      = 4
    columnImgId     = 5
    actDay          = 7
    formatAs        = "   -   "
    customerReturn  = 0
       
    today = datetime.datetime.today()
    thisHour = datetime.datetime.now().strftime('%I')    
    
    two_hours_ago = datetime.datetime.now() - datetime.timedelta(hours=2)
    d = modification_date(partialfilename+".csv")

    if d < two_hours_ago :
        return "                                                          Restart Linnworks ---------> " + partialfilename + ".csv is not current \n\n" 
    
    lines = ""

    partialfilename = "C:\myFolder\PosInv.csv"
    
    totValue = 0

    fib1  = open(partialfilename,encoding='UTF8')    
    fi_b1 = csv.reader(fib1)
    next (fi_b1)
    
    fob1  = open("C:\myFolder\Tmp.csv",'w')
    fo_b1 = csv.writer(fob1,lineterminator='\n')

    for row1 in fi_b1:
        departmnt = row1[0]
        totValue  = int(float(row1[1]))
        
        if departmnt == "TILE":
            fo_b1.writerow([departmnt,totValue,200000,totValue-200000])
        elif departmnt == "LAMINATE":
            fo_b1.writerow([departmnt,totValue,65000,totValue-65000])
                
    fib1.close()
    fob1.close()

    sortLinesInFile("C:\myFolder\Tmp.csv")
    
    f = open("C:\myFolder\Tmp.csv", "r")
    for i, line in enumerate(f):
        lines = lines + "        " + line.split(",")[0].zfill(15) + formatAs + line.split(",")[1].zfill(3) + formatAs + line.split(",")[2].zfill(3) + formatAs + line.split(",")[3].zfill(3) + "\n"
    f.close()
    
    return lines

    
def sortLinesInFile(fileName):
    f = open(fileName, "r")
    lines = [line for line in f if line.strip()]
    f.close()
    lines.sort()
       
    f = open(fileName, 'w')
    f.writelines(lines)
    f.close()    

def modification_date(filename):
    t = os.path.getmtime(filename)
    return datetime.datetime.fromtimestamp(t)
    
def constructEmail(emailSubjectOrBody,ordersDataList):   
    today    = datetime.date.today().strftime('%x')
    thisHour = datetime.datetime.now().strftime('%I:%M:%S %p')
    now_time = datetime.datetime.now()
    
    if   emailSubjectOrBody == "S" and now_time > now_time.replace(hour=9, minute=0, second=0, microsecond=0) :
        return "Inventory Targets : " + str(today) + "  " + str(thisHour) 
               
    elif emailSubjectOrBody == "B" and now_time > now_time.replace(hour=9, minute=0, second=0, microsecond=0) :
        body =        "PAYLESS COMPONENTS : \n******************\n" 
        body = body + "        Listings by the Hour today: \n" + ordersDataList[0] + "\n"
        return body 
    
def sendEmail(emailSubject,emailBody):
    me = "reports@paylesscomponents.com"
    you = "don@paylesscomponents.com"
    #you = ["don@paylesscomponents.com", "sales@paylesscomponents.com"]
    #you = ["kyle@paylesscomponents.com", "don@paylesscomponents.com", "nick@paylesscomponents.com"]
    #you = ["kyle@paylesscomponents.com", "jhamamy@factory-surplus.com", "nick@paylesscomponents.com", "desirae@paylesscomponents.com", "don@paylesscomponents.com"]
    
    COMMASPACE = ', '

    msg = MIMEMultipart('alternative')
    msg['Subject'] = emailSubject
    msg['From'] = me
    msg['To'] = COMMASPACE.join(you)

    part1 = MIMEText(emailBody, 'plain')       
    msg.attach(part1)
    
    s = smtplib.SMTP_SSL('smtpout.secureserver.net',465)       
    s.login("don@paylesscomponents.com", "donpay123")
    
    s.sendmail(me, you, msg.as_string())
    s.quit()

#cleanData()

processData()