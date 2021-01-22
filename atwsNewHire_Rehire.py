#daily hire/rehire

import atws
import atws.monkeypatch.attributes
import datetime
import pandas as pd
import numpy as np
import time
#pip install xlrd for excel files

#EV5 store code: Autotask customer ID
stores = {
        "ABC":   175,
        "ACJ":   200,
        "AOP":   222,
        "BHP":   223,
        "BRC":   224,
        "CBF":   225,
        "CBH":   188,
        "CBK":   189,
        "CBT":   226,
        "CBV":   227,
        "CNC":   228,
        "CNT":   229,
        "CRC":   190,
        "CRI":   191,
        "CRN":   193,
        "CRT":   194,
        "DMC":   230,
        "DMT":   231,
        "DRH":   232,
        "DVN":   198,
        "DVW":   199,
        "FFC":   233,
        "FOS":   201,
        "GPF":   234,
        "GVF":   235,
        "GWF":   202,
        "GWN":   203,
        "HFT":   236,
        "HSK":   204,
        "HSK":   205,
        "ICB":   177,
        "IOP":   208,
        "IOS":   206,
        "JMF":   237,
        "JMM":   238,
        "JMT":   239,
        "KKC":   240,
        "KKL":   210,
        "KKT":   211,
        "LSG":   241,
        "MAT":   196,
        "MCA":   242,
        "MGF":   243,
        "MLS":   195,
        "NMT":   244,
        "NPT":   245,
        "PNA":   207,
        "PRF":   246,
        "PRN":   209,
        "RCA":   212,
        "RCJ":   247,
        "REL":   248,
        "RIS":   249,
        "RNA":   213,
        "RSC":   250,
        "RTB":   214,
        "RUA":   215,
        "SCL":   251,
        "SFA":   252,
        "SFC":   253,
        "SHW":   254,
        "SLA":   255,
        "STN":   256,
        "TDB":   257,
        "TEF":   258,
        "TNG":   216,
        "TOD":   259,
        "TOR":   260,
        "TRN":   261,
        "TTG":   217,
        "VAN":   262,
        "VGA":   263,
        "VGC":   220,
        "VGH":   218,
        "VGT":   219,
        "VGY":   221,
        "VHT":   264,
        "VKC":   265,
        "VOD":   192,
        "VOP":   197,
        "VPT":   266,
        "WFT":   267
}
#autotask API connection
at = atws.connect(username='apiuser@domain.com',
                  password='--',
                  apiversion=1.6,
                  integrationcode='--')

#pandas data file read
data = pd.read_excel("./Daily Hire Rehire Report-Service Desk V2 3.12.2020.xls")
#print(data['Job Code'])

for index, row in data.iterrows():
        ticket = at.new('Ticket')
        if(pd.isnull(row["Rehire Date"])):
            ticket.Title = 'SAR New Hire: '+str(row["Name"])
        else:
           ticket.Title = 'SAR Re-Hire: '+row["Name"]+' '+str(row["Rehire Date"])
        ticket.Description = 'Name: '+str(row["Name"])+'\n'+'EmpleID: '+str(row["Empl ID"])+'\n'+'Action Date: '+str(row["Action Date"])+'\n'+'Job Code: '+row["Job Code"]+'\n'+'Job Title: '+row["Job Title"]
                                
        ticket.Source = at.picklist['Ticket']['Source']['Internal']
        ticket.IssueType = at.picklist['Ticket']['IssueType']['User Access']
        if row["Company"] in stores:
            ticket.AccountID = stores[row["Company"]]
        ticket.DueDateTime = datetime.datetime.now() + datetime.timedelta(days=2)
        ticket.Priority = at.picklist['Ticket']['Priority']['Low']
        ticket.Status = at.picklist['Ticket']['Status']['New']
        ticket.QueueID = at.picklist['Ticket']['QueueID']['SAR']
        #if you are just submitting one ticket:
        ticket.create() # updates the ticket object inline using CRUD patch
        print(ticket)
        time.sleep(.5)
        #print(at.picklist['Ticket']['QueueID'])



