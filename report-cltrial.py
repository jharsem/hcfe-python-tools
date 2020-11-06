import os
import pyodbc 
import datetime

import csv
import report_config as config
import report_email as email

# Build CSV
def buildCSV(fr,to,title,cursor_data):
    filename = "report-"+title+"-"+fr+"-to-"+to+".csv"
    with open(filename, "w", newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow([i[0] for i in cursor.description]) # write headers
        csv_writer.writerows(cursor_data)
    return filename
    
# Global

# Get date ranges
today = datetime.date.today()
dt_from = (today - datetime.timedelta(days=14))
dt_to = (today - datetime.timedelta(days=7))

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
select co.USER_ADDED, p.ID, co.BRANCH_IDENTIFIER
from CONTACT_LENS_ORDERS co with (NOLOCK), patient p with (NOLOCK) 
where co.trial_only = 1 
and p.id = co.PATIENTID
and co.FOLLOWED_UP_DATE is null
and co.COLLECTED_DATE between '{}' and '{}'
order by co.BRANCH_IDENTIFIER asc""".format(str_from, str_to)

# Run Query
trials = cursor.execute(str_query) 

print ("------------")
filename = buildCSV(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"),'trialsrep',trials)
descr = """
    Your report is based on the date range from {} to {}
    
    Report Notes
     
    This report shows collected trials that haven't been followed up for the time period stated above"""
subject = 'HCFE Trial Collection Report for {} to {}'
to = 'whitneyl@hcfeye.com.au,jonh@hcfeye.com.au' 
email.sendReport(filename,to,subject,descr,dt_to,dt_from)
os.rename(filename, './data/'+filename)
