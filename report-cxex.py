import os
import pyodbc 
import datetime

import csv
import report_config as config
import report_email as email

branches = ['BLA','BON','BRO','CHA','HUR','PAR','SYD']
tmp = { 
        "BLA":{
            "Cx":0,
            "Ex":0
            },
        "BON":{
            "Cx":0,
            "Ex":0
            },
        "BRO":{
            "Cx":0,
            "Ex":0
            },
        "CHA":{
            "Cx":0,
            "Ex":0
            },
        "HUR":{
            "Cx":0,
            "Ex":0
            },
        "PAR":{
            "Cx":0,
            "Ex":0
            },
        "SYD":{
            "Cx":0,
            "Ex":0
            }
    }

def findInValueTupleArray(needle, haystack):
    for row in haystack:
        if (row[1] == needle):
            return row[0]
    return 0

def computeArr ( buy, ret, exams ):
    for row in buy:
        tmp[row[1]]["Cx"] = row[0] - findInValueTupleArray(row[1],ret)
        tmp[row[1]]["Ex"] = findInValueTupleArray(row[1],exams)
       
# Build CSV
def buildCSV(fr,to,title,data):
    filename = "./data/report-"+title+"-"+fr+"-to-"+to+".csv"
    with open(filename, "w", newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Branch','Glasses Cx','Glasses/Lenses Exams']) # write headers
        for branch in data:
            csv_writer.writerow([branch, data[branch]['Cx'], data[branch]['Ex']])
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
SELECT COUNT(DISTINCT(i.PATIENTID)) as CXNUM, i.BRANCH_IDENTIFIER
FROM dbo.INVOICE i with (NOLOCK), dbo.INVOICE_ITEMS ii with (NOLOCK)
WHERE ii.INVOICEID = i.ID
AND i.SALE_DATE BETWEEN '""" + str_from + "' AND '"+ str_to + """' 
AND ii.STOCK_TYPE in (2,3,4)
AND i.TYPE = 2
AND i.DESPENSE_TOTAL > 0
GROUP BY i.BRANCH_IDENTIFIER"""

# Run Query
cx_buy = cursor.execute(str_query).fetchall() 

# Build SQL string for Return customers
str_query = """
SELECT COUNT(DISTINCT(i.PATIENTID)) as CXNUM, i.BRANCH_IDENTIFIER
FROM dbo.INVOICE i with (NOLOCK), dbo.INVOICE_ITEMS ii with (NOLOCK)
WHERE ii.INVOICEID = i.ID
AND i.SALE_DATE BETWEEN '"""+ str_from + "' AND '" + str_to + """' 
AND ii.STOCK_TYPE in (2,3,4)
AND i.TYPE = 6
AND i.DESPENSE_TOTAL < 0
GROUP BY i.BRANCH_IDENTIFIER"""

# Run Query
cx_return = cursor.execute(str_query).fetchall() 

# Build total visitor SQL string
'''
This is a measure of unique members with appointments that could lead to potential glasses/lenses sales
'''

str_query = """
Select count(distinct (a.PATIENTID)), a.BRANCH_IDENTIFIER
FROM APPOINTMENT a with (NOLOCK)
WHERE a.STARTDATE BETWEEN '"""+ str_from + "' AND '"+ str_to + """' 
AND a.APPOINTMENT_TYPE in ('CD','CF','CFD','DB','DBV','IC','ICY','LC','LCD','LCV','LDB','LDF','SC','SCY','SDF')
AND a.APP_PROGRESS = 5 
GROUP BY a.BRANCH_IDENTIFIER
"""

cx_exams = cursor.execute(str_query).fetchall()

# Process Rows.
computeArr(cx_buy,cx_return, cx_exams)
print ("------------")
filename = buildCSV(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"),'weeklyce',tmp)
descr = """
     Your report is based on the date range from {} to {}
    
     Report Notes
     
         Glasses Cx: A count of invoices with speclens items less returns and warranties. 
         Glasses/Lenses Exams: is a measure of appointment types which could lead to glasses sales 

         Both are counted within the above timeframe."""
subject = "HCFE Cx/Ex Report for {} to {}"
to = 'admin@hcfeye.com.au' 
email.sendReport(filename,to,subject,descr,dt_to,dt_from)
os.rename(filename, './data/'+filename)
