# Optom FTE Report 
# Author Jon <jharsem@hcf.com.au>
# Date: 2020-09
# Report produces daily forward looking FTE counts 

import os
import pyodbc 
import datetime
import smtplib
import calendar
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# For guessing MIME type based on file name extension
import mimetypes
from email.message import EmailMessage

import report_config as config
import report_email as email  

def buildArr (data_t):
    tmp = { 
        "BLA":{
            "data":0
            },
        "BON":{
            "data":0
            },
        "BRO":{
            "data":0
            },
        "CHA":{
            "data":0
            },
        "HUR":{
            "data":0
            },
        "PAR":{
            "data":0
            },
        "SYD":{
            "data":0
            }
    }
    for row in data_t:
        tmp[row[0]]['data']=row[1]
    return tmp

def sectionStart(sn):
    gap = 2
    height = 9
    init_offset = 5
    return (sn*(height+gap))+init_offset-1

def repOrder(code):
    return config.REPORTS["FTE"]["order"].index(code)

# 7 = 0, 8 = 1, 9 = 2, 10 = 3 , 11 = 4 , 12 = 5,  1 = 6 , 2 = 7
def adjustedMonthIndex(day):
    return (day.month + 5) % 12


# Get the data and write it to the relevant column
def processVol(wb,fr,look_ahead_count):
    row = sectionStart(repOrder('PXVOL'))
    col = 3
    writeHeader(wb, 'Patient Volumes', row, False, fr, look_ahead_count)

    for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
        if (single_date.weekday() == 6):
            continue
        # Build total utilisation SQL string
        str_from = single_date.strftime("%Y-%m-%d %H:%M:%S")
        str_to = single_date.strftime("%Y-%m-%d 23:59:59")

        str_query = buildVolQuery(str_from,str_to)
        ftes = cursor.execute(str_query).fetchall()
        data = buildArr(ftes)
        # Write Data Section
        writePxVolumeCol(wb,row,col,data)
        col = col + 1

def processFTE(wb,fr,look_ahead_count):
    row = sectionStart(repOrder("FTE"))
    col = 3
    writeHeader(wb, 'FTE', row, True, fr, look_ahead_count)

    for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
        if (single_date.weekday() == 6):
            continue
        # Build total utilisation SQL string
        str_from = single_date.strftime("%Y-%m-%d %H:%M:%S")
        str_to = single_date.strftime("%Y-%m-%d 23:59:59")

        str_query = buildTimeQuery(str_from,str_to)
        ftes = cursor.execute(str_query).fetchall()
        data = buildArr(ftes)
        # Write Data Section
        writeFTECol(wb,row,col,data)
        col = col + 1

def processOC(wb,fr,look_ahead_count):
    col = 0
    oc = [[0 for i in range(21)] for j in range(9)]
    for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
        sum = 0
        row = 0
        if (single_date.weekday() == 6):
            continue
        # Build total utilisation SQL string
        str_from = single_date.strftime("%Y-%m-%d %H:%M:%S")
        str_to = single_date.strftime("%Y-%m-%d 23:59:59")

        str_query = buildOCQuery(str_from,str_to)
        ftes = cursor.execute(str_query).fetchall()
        data = buildArr(ftes)
        # Write Data Section
        for k in config.BRANCHES.keys():
            oc[row][col] = data[k]['data']
            sum = sum + data[k]['data']
            row = row + 1
        oc[row][col] = sum
        col = col + 1
    return oc

def writeHeaderAndStatic(wb,title,fr,to):
    fmt_title= wb.add_format({'bold': True,'font_size':'14'})
    fmt_center = wb.add_format({'align': 'center'})

    ws = wb.worksheets()[0]
    ws.write('A1', 'Optom FTE', fmt_title)
    ws.write('A2', 'From')
    ws.write('B2', fr.strftime('%Y-%m-%d'))
    ws.write('C2', 'To',fmt_center)
    ws.write('D2', to.strftime('%Y-%m-%d'))
    return wb 

def writeHeader(wb, title, row, budget_req, fr, look_ahead_count):
    fmt_bold = wb.add_format({'bold': True})

    ws = wb.worksheets()[0]
    col = 0
    rowcpy = row + 1
    ws.write(row,col,title)

    for k in config.BRANCHES.keys():
        ws.write(rowcpy,col,k,fmt_bold)
        if (budget_req == True):
            ws.write(rowcpy,col+1,config.BRANCHES[k]['o_fte_100'])
            ws.write(rowcpy,col+2,config.BRANCHES[k]['o_fte_85'])    
        rowcpy = rowcpy + 1
    ws.write(rowcpy,col,'Total')

        # Do we need the budget columns
    if (budget_req == True):        
        ws.write(row,col+1,'Budget (100%)')
        ws.write(row,col+2,'Budget (85%)')    
        summation = '=sum('+xl_rowcol_to_cell(row+1, 1)+':'+xl_rowcol_to_cell(rowcpy-1, 1)+')'
        ws.write(rowcpy,1,summation)
        summation = '=sum('+xl_rowcol_to_cell(row+1, 2)+':'+xl_rowcol_to_cell(rowcpy-1, 2)+')'
        ws.write(rowcpy,2,summation)
    
    col = 3
    for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
        if (single_date.weekday() == 6):
            continue
        ws.write(row,col,single_date.strftime("%d-%m-%Y"))
        col = col + 1
    
    #format columns     
    ws.set_column(1, col+2, 10)
    
    
def writeFTECol(wb, row, col, data):
    fmt_bold_num = wb.add_format({'bold': True, 'num_format' : '0.00'})
    fmt_num = wb.add_format({'num_format': '0.00'})
    ws = wb.worksheets()[0]
    row = row + 1
    start_row = row
    for k in config.BRANCHES.keys():
        ws.write(row,col,data[k]['data']/(7*60+0.75*60),fmt_num)
        row = row + 1
    summation = '=sum('+xl_rowcol_to_cell(start_row, col)+':'+xl_rowcol_to_cell(row-1, col)+')'
    ws.write(row,col,summation,fmt_bold_num)
    return wb    
    
    
def writePercentFTE(wb,fr, look_ahead_count):
    fmt_pnum = wb.add_format({'num_format': '0.0%'})
    ws = wb.worksheets()[0]

    row = sectionStart(repOrder("FTEP"))
    fte_row = sectionStart(repOrder("FTE"))+1

    writeHeader(wb, 'FTE %', row, True, fr, look_ahead_count)  
    row = row + 1
    budget_col = 2
    
    for k in config.BRANCHES.keys():
        col = 3
        for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
            if (single_date.weekday() == 6):
                continue
            formula = "="+ xl_rowcol_to_cell(fte_row, col)+"/"+xl_rowcol_to_cell(row, budget_col)+""
            ws.write(row,col,formula,fmt_pnum)
            col = col + 1    
        row = row + 1
        fte_row = fte_row + 1
    # write summation
    col = 3
    for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
        if (single_date.weekday() == 6):
                continue
        form = '='+xl_rowcol_to_cell(fte_row, col)+'/'+xl_rowcol_to_cell(fte_row, 2)
        ws.write(row,col,form,fmt_pnum)
        col = col + 1

def writeFTEAP(wb, fr, look_ahead_count, oc):
    fmt_pnum = wb.add_format({'num_format': '0.0%'})
    ws = wb.worksheets()[0]

    row = sectionStart(repOrder("FTEAP"))
    fte_row = sectionStart(repOrder("FTE"))+1

    writeHeader(wb, 'FTE Actual %', row, True, fr, look_ahead_count)  
    row = row + 1
    oc_row_offset = row

    for k in config.BRANCHES.keys():
        col = 3
        for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
            if (single_date.weekday() == 6):
                continue
            if (oc[row-oc_row_offset][col-3] != 0):
                formula = "="+ xl_rowcol_to_cell(fte_row, col)+"/"+str(oc[row-oc_row_offset][col-3])+""
            else:
                formula = "=0"
            ws.write(row,col,formula,fmt_pnum)
            col = col + 1    
        row = row + 1
        fte_row = fte_row + 1
    # write summation
    col = 3
    for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
        if (single_date.weekday() == 6):
                continue
        form = '='+xl_rowcol_to_cell(fte_row, col)+'/'+str(oc[7][col-3])
        ws.write(row,col,form,fmt_pnum)
        col = col + 1


def writePxVolumeCol(wb, row, col, data):
    fmt_bold = wb.add_format({'bold': True})
    ws = wb.worksheets()[0]
    row = row + 1
    start_row = row
    for k in config.BRANCHES.keys():
        ws.write(row,col,data[k]['data'])
        row = row + 1
    summation = '=sum('+xl_rowcol_to_cell(start_row, col)+':'+xl_rowcol_to_cell(row-1, col)+')'
    ws.write(row,col,summation,fmt_bold)
    return wb    

#
#   Write the PX volume budget as a percentage value
#   incl helper function. 
#
def totalBudgetApptPerDay(day):
    rep_index = adjustedMonthIndex(day)
    sum = 0
    for k in config.BRANCHES.keys():
        sum = sum + config.REPORTS['APPT_PER_DAY'][k][rep_index]
    return sum
        

def writePXVolumePercentBudget(wb,fr, look_ahead_count):
    fmt_pnum = wb.add_format({'num_format': '0.0%'})
    ws = wb.worksheets()[0]

    row = sectionStart(repOrder("PXVOLP"))
    pxv_row = sectionStart(repOrder("PXVOL"))+1
    
    writeHeader(wb, 'Pt vols % of budget', row, False, fr, look_ahead_count)  
    ws.write(row,1,'Pts/day=')
    ws.write(row,2,'Dynamic')
    
    row = row + 1

    for k in config.BRANCHES.keys():
        col = 3
        for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
            if (single_date.weekday() == 6):
                continue
            rep_index = adjustedMonthIndex(single_date)
            #print(k,config.REPORTS['APPT_PER_DAY'][k][rep_index])
            formula = "="+ xl_rowcol_to_cell(pxv_row, col) + "/" + str(config.REPORTS['APPT_PER_DAY'][k][rep_index])
            ws.write(row,col,formula,fmt_pnum)
            col = col + 1    
        row = row + 1
        pxv_row = pxv_row + 1
    # write summation
    col = 3
    for single_date in (fr + datetime.timedelta(n) for n in range(look_ahead_count)):
        if (single_date.weekday() == 6):
                continue
        form = '=' + xl_rowcol_to_cell(pxv_row, col)+'/'+ str(totalBudgetApptPerDay(single_date))
        ws.write(row,col,form,fmt_pnum)
        col = col + 1



# Build XSLX file 
def buildXLSX(title,look_ahead_count):
    fr = datetime.date.today()
    to = fr + datetime.timedelta(look_ahead_count)
    fileName = './data/'+fr.strftime("%Y%m%d")+'-'+to.strftime('%Y%m%d')+'-optom-fte-report.xlsx'
    # Create a workbook and add a worksheet.
    wb = xlsxwriter.Workbook(fileName)
    wb.add_worksheet()
    
    writeHeaderAndStatic(wb,title, fr, to)
    processVol(wb, fr,look_ahead_count)
    processFTE(wb, fr,look_ahead_count)
    writePercentFTE(wb, fr, look_ahead_count)
    writePXVolumePercentBudget(wb,fr, look_ahead_count)
    oc = processOC(wb, fr, look_ahead_count)
    writeFTEAP(wb, fr, look_ahead_count, oc)
    wb.close()
    return fileName

#
# Query Strings
#
def buildVolQuery(str_from,str_to):
    str_query = """
    SELECT a.BRANCH_IDENTIFIER, count(a.ID)
    FROM dbo.APPOINTMENT a with (NOLOCK)
    WHERE a.STARTDATE between '""" + str_from + """' and '""" + str_to + """'
    AND a.APPOINTMENT_TYPE not in ('SP','LUN','BRK')
    GROUP BY a.BRANCH_IDENTIFIER
    """
    return str_query

def buildTimeQuery(str_from,str_to):
    str_query = """
    SELECT a.BRANCH_IDENTIFIER, sum(DATEDIFF(minute,a.STARTDATE,a.ENDDATE))
    FROM dbo.APPOINTMENT a with (NOLOCK)
    WHERE a.STARTDATE between '""" + str_from + """' and '""" + str_to + """'
    AND a.APPOINTMENT_TYPE not in ('SP','LUN','BRK')
    GROUP BY a.BRANCH_IDENTIFIER
    """
    return str_query

def buildOCQuery(str_from,str_to):
    str_query = """
    select a.BRANCH_IDENTIFIER, count(distinct(a.RESID))
    from dbo.APPOINTMENT a with (NOLOCK)
    WHERE a.STARTDATE between '""" + str_from + """' and '""" + str_to + """'
    and a.RESID in (
        Select ap.RESID
        from dbo.APPOINTMENT ap with (NOLOCK)
        WHERE ap.STARTDATE between '""" + str_from + """' and '""" + str_to + """'
        and ap.APPOINTMENT_TYPE not in ('SP','LUN','BRK')
    )
    group by a.BRANCH_IDENTIFIER
    """
    return str_query

# Global ########################################################################
look_ahead_count = 21 #three weeks
weekDays = ("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
#
today = datetime.date.today()

# SQL Connect Info
server = config.MSSQL_DB['server']
database = config.MSSQL_DB['database']
username = config.MSSQL_DB['username']
password = config.MSSQL_DB['password'] 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Build the XLSX file 
filename = buildXLSX("Optom FTE",look_ahead_count)
today = datetime.date.today()
dt_to = today + datetime.timedelta(look_ahead_count)

print ("------------")
descr = """
    Your report is based on the date range from {} to {}
    
    Report Notes
     
        This file is based on a 21 day lookahead and is only accurate on the day 
        it was generated as staffing and appointments will vary
    
    """
subject = "HCFE Optom FTE Report for {} to {}"
to = 'krisd@hcfeye.com.au,jonh@hcfeye.com.au' 
email.sendReport(filename,to,subject,descr,dt_to,today)