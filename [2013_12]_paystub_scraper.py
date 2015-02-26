#!/usr/bin/env python
#---------------------------------#
#     Paystub Data Scraper
#   Created by: Michael Magyar
#   Last Modified: 12/29/2013
#---------------------------------#
#
# Script will extract NRC paystub data (.txt) and dump to a (.csv) file format
# to perform analysis on.  (.xlsx) format is also possible, just uncomment
# applicable section.

import csv
import os
import re

##from openpyxl import Workbook
##from openpyxl.cell import get_column_letter

os.chdir(r'E:\Personal\Financial\Pay Stubs\Text')

Identity = ['','PP','Pay Period End','Pay Date','Gross PP Pay','Net PP Pay',
            'Federal Tax','State Tax','SS Tax','Medicare Tax',
            'FERS Retirement','Health Benefits','TSP','TSP Basic',
            'TSP Matching','TSP Total']

# Function to extract pertinent paystub data
def scrape(filelist):
    PayPeriod = []
    PayPeriodEndDate = []
    PayDate = []
    GrossPP = []
    NetPP = []
    FederalTax = []
    StateTax = []
    OASDITax = []
    MedicareTax = []
    FERS = []
    Health = []
    TSP = []
    TSPBasic = []
    TSPMatching = []
    TSPTotal = []
    CompleteList = []

    for infile in filelist:
        filelist.sort()
        with open(infile, 'r') as paystub:
            for line in paystub:
                #Pay Period
                if re.search(r'Pay Period #' ,line):
                    PP = re.search(r'(\d+)', line)
                    PayPeriod.append(int(PP.group(1)))
                #Pay Period End Date
                if re.search(r'Pay Period Ending' ,line):
                    PPD = re.search(r'(\d+.\d+.\d+)', line)
                    PayPeriodEndDate.append(PPD.group(1))
                #PayDate
                if re.search(r'Date of Paycheck' ,line):
                    PD = re.search(r'(\d+.\d+.\d+)', line)
                    PayDate.append(PD.group(1))
                #Gross PP
                if re.search(r'Gross Current' ,line):
                    GPP = re.search(r'(\d+.\d+)', line)
                    GrossPP.append(float(GPP.group(1)))
                #Net PP
                if re.search(r'Net Pay Current' ,line):
                    NPP = re.search(r'(\d+.\d+)', line)
                    NetPP.append(float(NPP.group(1)))
                #Federal Tax
                if re.search(r'Federal Taxes' ,line):
                    FTax = re.match(r'''\s+\w+\s\w+\s(\d+.\d+)\s+\w+\s+\w+\s+
                                    (\d+.\d+)|\s+\w+\s\w+\s+\w+\s+\w+\s+
                                    (\d+.\d+)''', line, re.X)
                    if FTax.group(1) and FTax.group(2):
                        FederalTax.append((float(FTax.group(1)) + \
                        float(FTax.group(2))))
                    else:
                        FederalTax.append(float(FTax.group(3)))
                #State Tax
                if re.search(r'State Tax 1' ,line):
                    STax = re.match(r'''\s+\w+\s\w+\s\d\s.\s\w+\s(\d+.\d+)\s+
                                    \w+\s+\w+\s+(\d+.\d+)|\s+\w+\s\w+\s\d\s.\s
                                    \w+\s+\w+\s+\w+\s+(\d+.\d+)''', line, re.X)
                    if STax.group(1) and STax.group(2):
                        StateTax.append((float(STax.group(1)) + \
                        float(STax.group(2))))
                    else:
                        StateTax.append(float(STax.group(3)))
                #OASDI Tax
                if re.search(r'OASDI Tax' ,line):
                    OTax = re.match(r'''\s+\w+\s\w+\s(\d+.\d+)\s+\w+\s+\d.\d
                                    \s+\w+\s+(\d+.\d+)|\s+\w+\s\w+\s+\w+\s+\d.
                                    \d\s+\w+\s+(\d+.\d+)''', line, re.X)
                    if OTax.group(1) and OTax.group(2):
                        OASDITax.append((float(OTax.group(1)) + \
                        float(OTax.group(2))))
                    else:
                        OASDITax.append(float(OTax.group(3)))
                #Medicare Tax
                if re.search(r'Medicare Tax' ,line):
                    MTax = re.match(r'''\s+\w+\s\w+\s(\d+.\d+)\s+\w+\s+\d.\d+
                                    \s+\w+\s+(\d+.\d+)|\s+\w+\s\w+\s+\w+\s+\d.
                                    \d+\s+\w+\s+(\d+.\d+)''', line, re.X)
                    if MTax.group(1) and MTax.group(2):
                        MedicareTax.append((float(MTax.group(1)) + \
                        float(MTax.group(2))))
                    else:
                        MedicareTax.append(float(MTax.group(3)))
                #FERS Retirement (pre FY2015)
                if re.search(r'FERS Retirement-Deduction' ,line):
                    F = re.match(r'''\s+\w+\s\w+.\w+\s+\w+\s+\w+\s+
                                 (\d+.\d+)|\s+\w+\s\w+.\w+\s+\w+\s+
                                 .\d\s+\w+\s+(\d+.\d+)''', line, re.X)
                    if F.group(1):
                        FERS.append(float(F.group(1)))
                    else:
                        FERS.append(float(F.group(2)))
                #FERS Retirement (Post FY2015)
                if re.search(r'Retirement - FERS', line):
                    F = re.match(r'''\s+\w+\s.\s\w+\s+\w+\s+.\d+\s+\w+\s+
                                (\d+.\d+)''', line, re.X)
                    FERS.append(float(F.group(1)))
                #Health
                if re.search(r'Health Benefits' ,line):
                    H = re.match(r'''\s+\w+\s\w+\s.\s\w+\s+\w+\s+\d+\s+\w+\s+
                                 (\d+.\d+)''', line, re.X)
                    Health.append(float(H.group(1)))
                #TSP (pre 2nd qtr 2012)
                if re.search(r'Thrift', line):
                    T = re.search(r'(\d+.\d+)', line)
                    TSP.append(float(T.group(1)))
                #TSP (current)
                if re.search(r'TSP Tax Deferred   Adjusted', line):
                    T = re.search(r'(\d+.\d+)', line)
                    TSP.append(float(T.group(1)))
                #TSP Basic
                if re.search(r'TSP Basic' ,line):
                    TB = re.search(r'(\d+.\d+)', line)
                    TSPBasic.append(float(TB.group(1)))
                #TSP Matching
                if re.search(r'TSP Matching' ,line):
                    TM = re.search(r'(\d+.\d+)', line)
                    TSPMatching.append(float(TM.group(1)))

        #Factors in no contributions or over-contributing
        if len(Health) != len(PayPeriod):
            Health.append(0)
        if len(TSP) != len(PayPeriod):
            TSP.append(0)
        if len(TSPMatching) != len(PayPeriod):
            TSPMatching.append(0)

    #Total TSP Contributions and Matching
    for i in range(0,len(TSP)):
        TSPTotal.append(TSP[i]+TSPBasic[i]+TSPMatching[i])

    #Compile all data to one list
    for i in range(0,len(PayPeriod)):
        CompleteList.append(['',PayPeriod[i],PayPeriodEndDate[i],PayDate[i],\
                             GrossPP[i],NetPP[i],FederalTax[i],StateTax[i],\
                             OASDITax[i],MedicareTax[i],FERS[i],Health[i],\
                             TSP[i],TSPBasic[i],TSPMatching[i],TSPTotal[i]])

    Column_Sum = ['','Total','','',sum(GrossPP),sum(NetPP),sum(FederalTax),\
                  sum(StateTax),sum(OASDITax),sum(MedicareTax),sum(FERS),\
                  sum(Health),sum(TSP),sum(TSPBasic),sum(TSPMatching),\
                  sum(TSPTotal)]

    return(CompleteList, Column_Sum)

os.chdir(r'E:\Personal\Financial\Pay Stubs\Extracted Data')
with open(r'compiled_data2014.csv', "wb") as ResultFile:
    count = int
    os.chdir(r'E:\Personal\Financial\Pay Stubs\Text')
    filelist = sorted(os.listdir(r'E:\Personal\Financial\Pay Stubs\Text'))
    Scraped, Totaled = scrape(filelist)

    #Write to .csv File
    wr = csv.writer(ResultFile, delimiter=',')
    wr.writerow(Identity)
    wr.writerows(Scraped)
    wr.writerow(Totaled)
    wr.writerow([])

# uncomment for (.xlxs) format
##os.chdir(r'Z:\Financial\Pay Stubs\Extracted Data')
##with open(r'compiled_data.csv', 'rU') as excel:
##    reader = csv.reader(excel, delimiter=',')
##
##    wb = Workbook()
##    dest_filename = r"compiled_data.xlsx"
##
##    ws = wb.worksheets[0]
##    ws.title = "Pay Stub Data"
##
##    for row_index, row in enumerate(reader):
##        for column_index, cell in enumerate(row):
##            column_letter = get_column_letter(column_index + 2)
##            ws.cell('%s%s'%(column_letter, (row_index + 5))).value = cell
##
##    wb.save(filename = dest_filename)
