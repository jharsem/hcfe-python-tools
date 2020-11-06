# Tool to clear forward Medicare Available > 3 days into future  
# Author Jon <jharsem@hcf.com.au>
# Date: 2020-11
# Unsets forward Medicare Eligibility > 3 days into future 

import os
import pyodbc 
import datetime
import calendar
import report_config as config
import logging
logging.basicConfig(filename='./logs/clear-eligibility.log', level=logging.DEBUG, format='%(asctime)s %(levelname)s %(message)s')

# Global ########################################################################

# Compute start date
today = datetime.date.today()
dt_to = today + datetime.timedelta(days=3)


# SQL Connect Info
server = config.MSSQL_DB['server']
database = config.MSSQL_DB['database']
username = config.MSSQL_DB['username']
password = config.MSSQL_DB['password'] 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Build total utilisation SQL string
str_to = dt_to.strftime("%Y-%m-%d 00:00:00")

str_query = """
    DELETE FROM dbo.PATIENT_CATEGORIES_LINK
    WHERE IDENTIFIER in ('ME','MNE') and PATIENT_ID in (
        select a.PATIENTID
        from dbo.APPOINTMENT a with (NOLOCK)
        Where a.STARTDATE > '{}')""".format(str_to)

total_rows = cursor.execute(str_query).rowcount
logging.info('Cleared %s Eligibility Rows',total_rows)
