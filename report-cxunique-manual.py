import os
import pyodbc 
import datetime
import calendar
import smtplib
# For guessing MIME type based on file name extension
import csv
import report_config as config
import report_email as email

branches = ['BLA','BON','BRO','CHA','HUR','PAR','SYD']

def computeArr ( rows, ex_list ):
    ex = [x[0] for x in ex_list]
    tmp = []
    print ("Exclusion Array Length:",len(ex))
    for row in rows:
        if row[0] not in ex:
            tmp.append(row)
    return tmp    

def buildArr (data_u):
    tmp = { 
        "BLA":{
            "unique":0,
            },
        "BON":{
            "unique":0,
            },
        "BRO":{
            "unique":0,
            },
        "CHA":{
            "unique":0,
            },
        "HUR":{
            "unique":0,
            },
        "PAR":{
            "unique":0,
            },
        "SYD":{
            "unique":0,
            }
    }
    for row in data_u:
        tmp[row[1]]["unique"]=tmp[row[1]]["unique"]+1
    return tmp
    
        
# Build CSV
def buildCSV(fr,to,title,data):
    filename = "report-"+title+"-"+fr+"-to-"+to+".csv"
    with open(filename, "w", newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Branch','Unique Yearly Cx FY20']) # write headers
        for branch in data:
            csv_writer.writerow([branch, data[branch]['unique']])
    return filename

# Global ########################################################################

# Get date ranges
dt_from = datetime.datetime(2019,10,1)
dt_to =  datetime.datetime(2019,10,31)

str_dt_from = dt_from.strftime("%Y-%m-%d %H:%M:%S")
str_dt_to = dt_to.strftime("%Y-%m-%d 23:59:59")


# SQL Connect Info
server = config.MSSQL_DB['server']
database = config.MSSQL_DB['database']
username = config.MSSQL_DB['username']
password = config.MSSQL_DB['password'] 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Build exclusion SQL string
str_query = """
SELECT DISTINCT(i.PATIENTID)
FROM dbo.INVOICE i with (NOLOCK)
WHERE i.SALE_DATE BETWEEN '2019-07-01 00:00:00' AND '{}'""".format(str_dt_from)
exclusion_list = cursor.execute(str_query).fetchall()

# Build SQL string
str_query = """
SELECT DISTINCT(i.PATIENTID) as NEWCx, p.BRANCH_IDENTIFIER
FROM dbo.INVOICE i with (NOLOCK), dbo.PATIENT p with (NOLOCK)
WHERE i.SALE_DATE BETWEEN '{}' AND '{}' 
AND i.PATIENTID = p.ID""".format(str_dt_from,str_dt_to)

# Run Query
rows = cursor.execute(str_query) 

# Process Rows.
unique_cx_rows = computeArr(rows,exclusion_list)
result = buildArr(unique_cx_rows)

print ("------------")
filename = buildCSV(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"),'monthly_unique_cx',result)
descr = """
    Your report is based on the date range from {} to {}
    
    Report Notes
     
        Unique Cx In FY: Counts of unique members with accounts created this Financial Year.
    
    """
subject = "HCFE Unique Cx In Year Report for {} to {}"
to = 'jonh@hcfeye.com.au' 
email.sendReport(filename,to,subject,descr,dt_to,dt_from)
os.rename(filename, './data/'+filename)
