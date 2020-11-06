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
import report_email as email

def computeArr ( rows, ex_list ):
    ex = [x[0] for x in ex_list]
    tmp = []
    print ("Exclusion Array Length:",len(ex))
    for row in rows:
        if row[0] not in ex:
            tmp.append(row)
    return tmp    

def buildArr (data_t):
    tmp = { 
        "BLA":{
            "total":0
            },
        "BON":{
            "total":0
            },
        "BRO":{
            "total":0
            },
        "CHA":{
            "total":0
            },
        "HUR":{
            "total":0
            },
        "PAR":{
            "total":0
            },
        "SYD":{
            "total":0
            }
    }
    for row in data_t:
        tmp[row[0]]["total"]=row[1]
    return tmp
    
        
# Build CSV
def buildCSV(fr,to,title,data):
    filename = "report-"+title+"-"+fr+"-to-"+to+".csv"
    with open(filename, "w", newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Branch','Utilisation Cx']) # write headers
        for branch in data:
            csv_writer.writerow([branch, data[branch]['total']])
    return filename

# Global ########################################################################

# Get date ranges
today = datetime.date.today()
dt_from = (today - datetime.timedelta(days=6))
dt_to = today

#dt_from = datetime.datetime(2020,6,1)
#dt_to =  datetime.datetime(2020,6,6)

str_dt_from =  dt_from.strftime("%Y-%m-%d 00:00:00")
str_dt_to = dt_to.strftime("%Y-%m-%d 23:59:59")

# SQL Connect Info
server = config.MSSQL_DB['server']
database = config.MSSQL_DB['database']
username = config.MSSQL_DB['username']
password = config.MSSQL_DB['password'] 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Build total utilisation SQL string
str_from = dt_from.strftime("%Y-%m-%d %H:%M:%S")
str_to = dt_to.strftime("%Y-%m-%d 23:59:59")
str_query = """
SELECT BRANCH_IDENTIFIER, sum(CXNUM)
FROM (
	SELECT COUNT(i.PATIENTID) AS "CXNUM", i.BRANCH_IDENTIFIER
	FROM dbo.INVOICE i with (NOLOCK)
	WHERE i.SALE_DATE between '""" + str_from + "' and '" + str_to + """'
	GROUP BY i.BRANCH_IDENTIFIER
    
UNION ALL
	
	SELECT COUNT(s.PATIENTID) AS "CXNUM",s.DELIVERY_BRANCH as BRANCH_IDENTIFIER
	FROM dbo.SPECTACLE_ORDERS s with (NOLOCK)
	WHERE s.COLLECTED_DATE between '"""+ str_from + "' and '" + str_to + """'
	GROUP BY s.DELIVERY_BRANCH

UNION ALL
	
	SELECT COUNT(c.PATIENTID) AS "CXNUM",c.BRANCH_IDENTIFIER
	FROM dbo.CONTACT_LENS_ORDERS c with (NOLOCK)
	WHERE c.COLLECTED_DATE between '""" + str_from + "' and '" + str_to + """'
	GROUP BY c.BRANCH_IDENTIFIER
) t
GROUP BY t.BRANCH_IDENTIFIER
"""

total_visitors = cursor.execute(str_query).fetchall()

print ("------------")
result = buildArr(total_visitors)
filename = buildCSV(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"),'weekly_utilisation',result)
descr = """
    Your report is based on the date range from {} to {}
    
    Report Notes
     
        Utilisation Cx: is a measure of all accounts created (including consultation), 
        and (gl/cl) deliveries done for a given timespan (non unique!)
    
    """
subject = "HCFE Utilisation Report for {} to {}"
to = 'admin@hcfeye.com.au' 
email.sendReport(filename,to,subject,descr,dt_to,dt_from)
os.rename(filename, './data/'+filename)