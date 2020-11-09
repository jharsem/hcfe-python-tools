import os
import pyodbc
import calendar
import datetime
import smtplib
# For guessing MIME type based on file name extension
import mimetypes
from email.message import EmailMessage
import csv
import report_config as config

branches = ['BLA','BON','BRO','CHA','HUR','PAR','SYD']
sql = ['appt_px_all','Gls Appt No','Cl Appt No','cx_all','Gl Cx No','Cl Cx No','CL Cx Per RX','Total Appt No','OCT No','CL New Fit','Cl Refit','Unique Cx in Month']

tmp = {}

# Get date ranges
today = datetime.date.today()
dt_from = (today - datetime.timedelta(days=6))
dt_to = today

str_dt_from =  dt_from.strftime("%Y-%m-%d 00:00:00")
str_dt_to = dt_to.strftime("%Y-%m-%d 23:59:59")

sqlstrings = {
    "Total Appt No":"""
        Select a.BRANCH_IDENTIFIER,count(a.ID) 
        FROM APPOINTMENT a with (NOLOCK)
        WHERE a.STARTDATE BETWEEN '{}' AND '{}' 
        AND a.APPOINTMENT_TYPE not in ('BRK','LUN','REP','SP')
        AND a.APP_PROGRESS = 5 
        GROUP BY a.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to),
    "appt_px_all":"""
        Select a.BRANCH_IDENTIFIER,count(distinct (a.PATIENTID)) 
        FROM APPOINTMENT a with (NOLOCK)
        WHERE a.STARTDATE BETWEEN '{}' AND '{}' 
        AND a.APPOINTMENT_TYPE not in ('BRK','LUN','REP','SP')
        AND a.APP_PROGRESS = 5 
        GROUP BY a.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to),
    "Cl Appt No":"""
        Select a.BRANCH_IDENTIFIER,count(distinct (a.PATIENTID)) 
        FROM APPOINTMENT a with (NOLOCK)
        WHERE a.STARTDATE BETWEEN '{}' AND '{}' 
        AND a.APPOINTMENT_TYPE in ('CDF','CL','CLD','CLV')
        AND a.APP_PROGRESS = 5 
        GROUP BY a.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to),
    "Gls Appt No":"""
        Select a.BRANCH_IDENTIFIER,count(distinct (a.PATIENTID)) 
        FROM APPOINTMENT a with (NOLOCK)
        WHERE a.STARTDATE BETWEEN '{}' AND '{}' 
        AND a.APPOINTMENT_TYPE in ('CD','CF','CFD','DB','DBV','IC','ICY','LC','LCD','LCV','LDB','LDF','SC','SCY','SDF')
        AND a.APP_PROGRESS = 5 
        GROUP BY a.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to),
    "cx_all":"""
        Select i.BRANCH_IDENTIFIER, count(Distinct(i.PATIENTID))
        FROM INVOICE i with (NOLOCK)
        WHERE i.SALE_DATE between '{}' and '{}' 
        AND i.DESPENSE_TOTAL >0
        AND i.[TYPE]=2
        group by i.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to),
    "Gl Cx No":"""
        Select i.BRANCH_IDENTIFIER, count(Distinct(i.PATIENTID))
        FROM INVOICE i with (NOLOCK), INVOICE_ITEMS ii with (NOLOCK)
        WHERE i.SALE_DATE between '{}' and '{}' 
        AND i.DESPENSE_TOTAL >0
        AND ii.INVOICEID = i.id
        AND i.[TYPE]=2
        AND ii.STOCK_TYPE in (2,3,4)
        group by i.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to),
    "Cl Cx No":"""
        Select i.BRANCH_IDENTIFIER, count(Distinct(i.PATIENTID))
        FROM INVOICE i with (NOLOCK), INVOICE_ITEMS ii with (NOLOCK)
        WHERE i.SALE_DATE between '{}' and '{}' 
        AND i.DESPENSE_TOTAL >0
        AND i.[TYPE]=2
        AND ii.INVOICEID = i.id
        AND ii.STOCK_TYPE in (5)
        group by i.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to), 
    "CL Cx Per RX":"""
        Select co.BRANCH_IDENTIFIER,count(distinct(co.PATIENTID))
        FROM dbo.CONTACT_LENS_ORDERS co with (NOLOCK)
        WHERE co.ORDER_DATE between '{}' and '{}'
        AND co.CANCELLED = 0
        AND co.TRIAL = 0
        AND co.WARRANTY = ''
        AND co.RX_REFERENCE in (SELECT cr.ID FROM dbo.CONTACT_RX cr with (NOLOCK) where cr.DATE_ADDED between DATEADD(MONTH,-4,co.ORDER_DATE) and co.ORDER_DATE)
        GROUP by co.BRANCH_IDENTIFIER 
        """.format(str_dt_from,str_dt_to),
    "OCT No":"""
        Select i.BRANCH_IDENTIFIER, count(Distinct(i.PATIENTID))
        FROM INVOICE i with (NOLOCK), INVOICE_ITEMS ii with (NOLOCK)
        WHERE i.SALE_DATE between '{}' and '{}' 
        AND i.[TYPE]=2
        AND ii.INVOICEID = i.id
        AND ii.ITEM_BARCODE = 'OCT'
        group by i.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to), 
    "CL New Fit":"""
        Select i.BRANCH_IDENTIFIER, count(Distinct(i.PATIENTID))
        FROM INVOICE i with (NOLOCK), INVOICE_ITEMS ii with (NOLOCK)
        WHERE i.SALE_DATE between '{}' and '{}' 
        AND i.[TYPE]=2
        AND ii.INVOICEID = i.id
        AND ii.ITEM_BARCODE = 'LTPRO'
        group by i.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to),
    "Cl Refit":"""
        Select a.BRANCH_IDENTIFIER, count(distinct (a.PATIENTID)) 
        FROM APPOINTMENT a with (NOLOCK)
        WHERE a.STARTDATE BETWEEN '{}' AND '{}' 
        AND a.APPOINTMENT_TYPE in ('AL','SL')
        AND a.APP_PROGRESS = 5 
        GROUP BY a.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to),
    "Unique Cx in Month":"""
        SELECT i.BRANCH_IDENTIFIER, count(DISTINCT(i.PATIENTID))
        FROM INVOICE i with (NOLOCK)
        WHERE i.SALE_DATE BETWEEN '{}' AND '{}' 
        GROUP BY i.BRANCH_IDENTIFIER""".format(str_dt_from,str_dt_to)
    }
def findValueForBranch(branch,data):
    for (b,d) in data:
        if (b == branch):
            return d
    return 0

# Build CSV
def buildCSV(fr,to,title):
    filename = "report-"+title+"-"+fr+"-to-"+to+".csv"
    with open(filename, "w", newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        header = ['branch']
        for s in sql:
            header.append(s)
        csv_writer.writerow(header) # write headers
        for branch in branches:
            row = [branch]
            for s in sql:
                row.append(findValueForBranch(branch,tmp[s]))
            csv_writer.writerow(row)
    return filename
    
# Send Email
def sendReport(f):
    fromaddr = "Report Robot <reports@hcfeye.com.au>"
    toaddrs  = "admin@hcfeye.com.au"
    #toaddrs  = "Teresa Ou <teresao@hcfeye.com.au>"
    msg = EmailMessage()
    messageText = """
    Your report is based on the date range from {} to {}
    
    Report Notes
     
        cx_all         # all non-warranty customers
        Cl Cx No       # all cl customers
        Gl Cx No       # all speclens customers
        CL Cx Per RX   # Cl orders with script < 4 months
        appt_px_all    # all exams
        Cl Appt No     # all cl exams (idea: cl appointments)
        Gls Appt No    # all glasses exams (idea: glasses appointments)
        OCT No         # OCT patients 
        CL New Fit     # New Fits (lens teaches) 
        Cl Refit       # Refit (based on appointments)
        Unique Cx in Month  # Unique Patients per Month (based in invoices, including Medicare) no exclusion

        NOTE: All of the above are given as unique patient count.

        Total Appt No  # all completed appointments (excluding rep, break, lunch and special) 
    
    """.format(str_dt_from, str_dt_to)
    msg.set_content(messageText)
    msg['Subject'] = "HCFE Weekly Base Number Report for {} to {}".format(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"))
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
for s in sql:
    result = cursor.execute(sqlstrings[s]).fetchall() 
    tmp[s] = result
# Process Rows.
print ("------------")
filename = buildCSV(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"),'masterdata')
sendReport(filename)
os.rename(filename, './data/'+filename)
