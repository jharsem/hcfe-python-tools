import os
import pyodbc 
import datetime

import csv
import report_config as config
import report_email as email

branches = ['BLA','BON','BRO','CHA','HUR','PAR','SYD']
tmp = { 
        "BLA":{
            "ClCx":0,
            "ClEx":0
            },
        "BON":{
            "ClCx":0,
            "ClEx":0
            },
        "BRO":{
            "ClCx":0,
            "ClEx":0
            },
        "CHA":{
            "ClCx":0,
            "ClEx":0
            },
        "HUR":{
            "ClCx":0,
            "ClEx":0
            },
        "PAR":{
            "ClCx":0,
            "ClEx":0
            },
        "SYD":{
            "ClCx":0,
            "ClEx":0
            }
    }

def findInValueTupleArray(needle, haystack):
    for row in haystack:
        if (row[1] == needle):
            return row[0]
    return 0

def computeArr ( buy, exams ):
    for row in buy:
        tmp[row[0]]["ClCx"] = row[1]
        tmp[row[0]]["ClEx"] = findInValueTupleArray(row[1],exams)
       
# Build CSV
def buildCSV(fr,to,title,data):
    filename = "./data/report-"+title+"-"+fr+"-to-"+to+".csv"
    with open(filename, "w", newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Branch','Cl Cx','Cl Exams']) # write headers
        for branch in data:
            csv_writer.writerow([branch, data[branch]['ClCx'], data[branch]['ClEx']])
    return filename

# Global

# Get date ranges
today = datetime.date.today()
dt_from = (today - datetime.timedelta(days=6))
dt_to = today

#dt_from = datetime.datetime(2019,10,1)
#dt_to =  datetime.datetime(2019,10,31)

str_from = dt_from.strftime("%Y-%m-%d 00:00:00")
str_to = dt_to.strftime("%Y-%m-%d 23:59:59")

# SQL Connect Info
server = config.MSSQL_DB['server']
database = config.MSSQL_DB['database']
username = config.MSSQL_DB['username']
password = config.MSSQL_DB['password'] 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Build SQL string for Purchase customers
str_query = """
    Select co.BRANCH_IDENTIFIER, count(co.ID)
    FROM dbo.CONTACT_LENS_ORDERS co with (NOLOCK)
    WHERE co.ORDER_DATE between '{}' and '{}'
    AND co.CANCELLED = 0
    AND co.TRIAL = 0
    AND co.WARRANTY = '' 
    AND co.RX_REFERENCE in (SELECT cr.ID FROM dbo.CONTACT_RX cr with (NOLOCK) 
        WHERE cr.DATE_ADDED between DATEADD(MONTH,-3,co.ORDER_DATE) AND co.ORDER_DATE)
    GROUP BY co.BRANCH_IDENTIFIER""".format(str_from,str_to)    

# Run Query
cl_rx_buy = cursor.execute(str_query).fetchall() 

# Build total visitor SQL string
'''
This is a measure of unique members with appointments that could lead to cl sales
'''

str_query = """
    Select count(distinct (a.PATIENTID)), a.BRANCH_IDENTIFIER
    FROM APPOINTMENT a with (NOLOCK)
    WHERE a.STARTDATE BETWEEN '"""+ str_from + "' AND '"+ str_to + """' 
    AND a.APPOINTMENT_TYPE in ('CDF','CL','CLD','CLV')
    AND a.APP_PROGRESS = 5 
    GROUP BY a.BRANCH_IDENTIFIER
    """

cl_exams = cursor.execute(str_query).fetchall()

# Process Rows.
computeArr(cl_rx_buy, cl_exams)
print ("------------")
filename = buildCSV(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"),'weekly_clcx_rx_vs_clex',tmp)
descr = """Your report is based on the date range from {} to {}
    
    Report Notes
     
        Contact Cl: A count of invoices with contact lens items (based on scripts in a 4 month window). 
        Cl Exams: is a measure of appointment types which could lead to CL sales 

        both are counted within the above timeframe."""
subject = 'HCFE Cx/Ex Report for {} to {}'
to = 'admin@hcfeye.com.au' 
email.sendReport(filename,to,subject,descr,dt_to,dt_from)
os.rename(filename, './data/'+filename)
