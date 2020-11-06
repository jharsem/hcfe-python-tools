# Optom Stats Report 
# Author Jon <jharsem@hcf.com.au>
# Date: 2020-07
# Report produces monthly optometrist statistics

import os
import pyodbc 
import datetime
import smtplib
import calendar
# For guessing MIME type based on file name extension
import mimetypes
from email.message import EmailMessage
import csv
import report_config as config

optometrists = []
reports = ['#cl_cx','$cl','#cl_appt','#lt_appt','#lens','$lens','#upx','#ex','$cons','#w_days','#910/11','#912/13/14','#916/18','#940/41','#tmpl','#amd','#cat','#lasik']
tmp = {}

# Get date ranges
year = datetime.date.today().year
previous_month = datetime.date.today().month - 1 or 12
dt_from = datetime.datetime(year,previous_month,1)
dt_to = datetime.datetime(year,previous_month,calendar.monthrange(year,previous_month)[1])

str_dt_from =  dt_from.strftime("%Y-%m-%d 00:00:00")
str_dt_to = dt_to.strftime("%Y-%m-%d 23:59:59")

# Report column definitions   
sqlstrings = {
    "#cl_cx":"""SELECT one.USER_IDENTIFIER, (ISNULL(one.sales,0) - ISNULL(two.sales,0))
        FROM (
            SELECT i.USER_IDENTIFIER, COUNT(DISTINCT i.PATIENTID) as sales
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN INVOICE_ITEMS ii on ii.INVOICEID = i.ID 
            WHERE SALE_DATE BETWEEN '{}' and '{}'
            and i.TYPE = 2
            and ii.STOCK_TYPE in (5)
            GROUP BY i.USER_IDENTIFIER
        ) AS one
        LEFT JOIN 
        (
            SELECT i.USER_IDENTIFIER, COUNT(DISTINCT i.PATIENTID) as sales
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN INVOICE_ITEMS ii on ii.INVOICEID = i.ID 
            WHERE SALE_DATE BETWEEN '{}' and '{}'
            and i.TYPE = 6
            and ii.STOCK_TYPE in (5)
            GROUP BY i.USER_IDENTIFIER
            ) AS two
        ON one.USER_IDENTIFIER = two.USER_IDENTIFIER""",
    "$cl":"""SELECT i.USER_IDENTIFIER, SUM(ii.QTY * ii.UNITPRICE - ii.DISCOUNT_AMOUNT) 
	    FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND ii.STOCK_TYPE = 5
            AND i.TYPE in (2,6)
	    GROUP BY i.USER_IDENTIFIER""",
    "#cl_appt":"""SELECT a.RESID, COUNT(a.ID) 
	    FROM dbo.APPOINTMENT a with (NOLOCK)
            WHERE a.STARTDATE BETWEEN '{}' AND '{}' 
            AND a.APPOINTMENT_TYPE in ('CDF','CL','CLD','CLV')
            AND a.APP_PROGRESS = 5 
        GROUP BY a.RESID""",
    "#lt_appt":"""SELECT a.RESID, COUNT(a.ID) 
	    FROM dbo.APPOINTMENT a with (NOLOCK)
            WHERE a.STARTDATE BETWEEN '{}' AND '{}' 
            AND a.APPOINTMENT_TYPE in ('LT','TDF','LTD','LTM','LTF')
            AND a.APP_PROGRESS = 5 
        GROUP BY a.RESID""",
    "#lens":"""SELECT i.USER_IDENTIFIER, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND ii.STOCK_TYPE = 4
            AND i.TYPE in (2,6)
            AND LEFT(ii.ITEM_BARCODE,3) <> 'GAP'
        GROUP BY i.USER_IDENTIFIER""",
    "$lens":"""SELECT i.USER_IDENTIFIER, SUM(ii.QTY * ii.UNITPRICE - ii.DISCOUNT_AMOUNT) 
	    FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND ii.STOCK_TYPE = 4
            AND i.TYPE in (2,6)
	    GROUP BY i.USER_IDENTIFIER""",
    "#upx":"""SELECT e.USER_IDENTIFIER, COUNT(DISTINCT e.PATIENT_ID) 
	    FROM dbo.EXAMINATION e with (NOLOCK)
        WHERE e.EXAM_DATE between '{}' and '{}'
	    GROUP BY e.USER_IDENTIFIER""",    
    "#ex":"""SELECT e.USER_IDENTIFIER, COUNT(e.ID) 
	    FROM dbo.EXAMINATION e with (NOLOCK)
        WHERE e.EXAM_DATE between '{}' and '{}'
	    GROUP BY e.USER_IDENTIFIER""",
    "$cons":"""SELECT i.USER_IDENTIFIER, SUM(ii.QTY * ii.UNITPRICE - ii.DISCOUNT_AMOUNT) 
	    FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.DESPENSE_TOTAL = 0
            AND i.ORDER_TYPE = 0
	    GROUP BY i.USER_IDENTIFIER""",
    "#w_days":"""SELECT cl.user_added, COUNT(cl.TIMESTMP) 
        FROM dbo.CLOCKINOUT cl with (NOLOCK)
        WHERE cl.TIMESTMP between '{}' and '{}'
        GROUP BY cl.USER_ADDED""",
    "#910/11":"""SELECT i.USER_IDENTIFIER, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND ii.ITEM_BARCODE in ('10900','10911')
        GROUP BY i.USER_IDENTIFIER""",                
    "#912/13/14":"""SELECT i.USER_IDENTIFIER, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND ii.ITEM_BARCODE in ('10912','10913','10914')
        GROUP BY i.USER_IDENTIFIER""",
    "#916/18":"""SELECT i.USER_IDENTIFIER, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND ii.ITEM_BARCODE in ('10916','10918')
        GROUP BY i.USER_IDENTIFIER""",
    "#940/41":"""SELECT i.USER_IDENTIFIER, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND ii.ITEM_BARCODE in ('10940','10941')
        GROUP BY i.USER_IDENTIFIER""",
    "#tmpl":"""SELECT l.USER_ADDED, count(l.ID)
        FROM dbo.PATIENT_LETTERS l with (NOLOCK)
            WHERE l.DATE_ADDED between '{}' and '{}'
            AND l.LETTER_DESCRIPTION in ('Treatment Plan','CL Treatment Plan')
        GROUP BY l.USER_ADDED""",
    "#amd":"""SELECT l.USER_ADDED, count(l.ID)
        FROM dbo.PATIENT_LETTERS l with (NOLOCK)
            WHERE l.DATE_ADDED between '{}' and '{}'
            AND l.LETTER_DESCRIPTION in ('Checklist - AMD Referral','AMD','Amd')
        GROUP BY l.USER_ADDED""",
    "#cat":"""SELECT l.USER_ADDED, count(l.ID)
        FROM dbo.PATIENT_LETTERS l with (NOLOCK)
            WHERE l.DATE_ADDED between '{}' and '{}'
            AND l.LETTER_DESCRIPTION in ('Checklist - Cataract Referral','CAT','Cat')
        GROUP BY l.USER_ADDED""",
    "#lasik":"""SELECT l.USER_ADDED, count(l.ID)
        FROM dbo.PATIENT_LETTERS l with (NOLOCK)
            WHERE l.DATE_ADDED between '{}' and '{}'
            AND l.LETTER_DESCRIPTION in ('Checklist - Lasik','LASIK','Lasik')
        GROUP BY l.USER_ADDED"""
    }

def findValueForOptometrist(disp,data):
    for (b,d) in data:
        if (b == disp):
            return d
    return 0

# Get all Dispensers
def getOptometrists():
    sql = """SELECT distinct(u.identifier) 
    FROM dbo.USERS u with (NOLOCK)
        WHERE u.user_type = 1
        AND u.inactive = 0
        AND u.IDENTIFIER not in ('','KTR','JMW','MIL','KWO','HTH')"""
    return cursor.execute(sql).fetchall() 

# Build CSV
def buildCSV(fr,to,title):
    filename = "report-"+title+"-"+fr+"-to-"+to+".csv"
    with open(filename, "w", newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        header = ['Optom Code']
        for r in reports:
            header.append(r)
        csv_writer.writerow(header) # write headers
        for optom in optometrists:
            row = [optom[0]]
            for r in reports:
                row.append(findValueForOptometrist(optom[0],tmp[r]))
            csv_writer.writerow(row)
    return filename


# Send Email
def sendReport(f):
    fromaddr = "Report Robot <reports@hcfeye.com.au>"
    #toaddrs  = "Jon Harsem <jonh@hcfeye.com.au>"
    toaddrs  = "Lauren Valentine <laurenv@hcfeye.com.au>"

    msg = EmailMessage()
    messageText = """
    Your report is based on the date range from {} to {}
    
    Report Notes

        #cl_cx      unique count CL patients (t.i.a returns)
        $cl         total cl sales $ (t.i.a returns)
        #cl_appt    Cl appointments (appt types: 'CDF','CL','CLD','CLV' )
        #lt_appt    Lens Teach appoitnemnts (appt types: 'LT','TDF','LTD','LTM','LTF')
        #lens       lenses sold (t.i.a returns)
        $lens       total lens sales (t.i.a returns)
        #upx        unique Px with exams
        #ex         total Px exams/consults
        $cons       total consult $ (t.i.a. reversals)
        #w_days     days worked 
        #910/11     10910 & 10911 billed (t.i.a. reversals)
        #912/13/14  10912, 10913 & 10914 billed (t.i.a. reversals)
        #916/18     10916 & 10918 billed (t.i.a. reversals)
        #940/41     10940 & 10941 billed (t.i.a. reversals)
        #trmpl      Treatment plans submitted
        #amd        AMD checklists submitted
        #cat        Cataract checklists submitted
        #lasik      Ophthal surgery checklists submitted
    
    """.format(str_dt_from, str_dt_to)
    msg.set_content(messageText)
    msg['Subject'] = "HCFE Optometrist Report for {} to {}".format(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"))
    msg['From'] = fromaddr
    msg['To'] = toaddrs
    # Send the message via our own SMTP server.
    s = smtplib.SMTP('192.168.2.41')
    path = os.path.join('.', f)
    ctype, encoding = mimetypes.guess_type(path)
    if ctype is None or encoding is not None:
        # No guess could be made, or the file is encoded (compressed), so
        # use a generic bag-of-bits type.
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    with open(path, 'rb') as fp:
        msg.add_attachment(fp.read(), maintype=maintype, subtype=subtype, filename=f)
    s.send_message(msg)
    s.quit()
    return

# Global

# SQL Connect Info
server = config.MSSQL_DB['server']
database = config.MSSQL_DB['database']
username = config.MSSQL_DB['username']
password = config.MSSQL_DB['password'] 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Build SQL string for Purchase customers
print ('Timeframe',str_dt_from,str_dt_to)
optometrists = getOptometrists()
for r in reports:
    if (sqlstrings[r].count('{}') == 2):
        sql = sqlstrings[r].format(str_dt_from,str_dt_to)
    if (sqlstrings[r].count('{}') == 4):
        sql = sqlstrings[r].format(str_dt_from,str_dt_to,str_dt_from,str_dt_to) 
    result = cursor.execute(sql).fetchall()
    tmp[r] = result
filename = buildCSV(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"),'optometrists')
# Process Rows.
print ("------------")
sendReport(filename)
os.rename(filename, './data/'+filename)
