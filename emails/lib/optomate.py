#   class Optomate
#   @author Jon Harsem <jonh@hcfeye.com.au>
#   @version 1.0

import os
import pyodbc 
import datetime
import email_config as config

class Optomate():
    """ This class handles the information to and from Optomate Touch """

    sqlstrings = {
        "feedback"  :   """
            SELECT a.ID, a.APPOINTMENT_TYPE, a.STARTDATE, a.BRANCH_IDENTIFIER as BRANCH, p.TITLE, p.ID as PID, HASHBYTES('SHA2_256',CAST (p.ID AS VARCHAR(max))) as GENID, p.GIVEN, 
                p.SURNAME, p.EMAIL, p.NOEMAIL, p.BIRTHDATE
            FROM dbo.APPOINTMENT a with (NOLOCK), dbo.PATIENT p with (NOLOCK)
            WHERE a.STARTDATE between '{start_date}' and '{end_date}' 
            AND a.APP_PROGRESS = 5 
            AND p.ID = a.PATIENTID
            AND p.ID not in (SELECT i.PATIENTID FROM invoice i with (NOLOCK) where i.SALE_DATE between '{start_date}' and '{end_date}' 
                AND i.DESPENSE_TOTAL > 0)""",
        "followup"  :   """
            SELECT p.TITLE, p.GIVEN, p.SURNAME, p.NOEMAIL, p.EMAIL, p.BIRTHDATE, s.BRANCH_IDENTIFIER as BRANCH 
            FROM dbo.SPECTACLE_ORDERS s with (NOLOCK), patient p with (nolock) 
            WHERE s.COLLECTED_DATE='{collected_date}' and p.ID = s.PATIENTID""",
        "peq"       :   """
            SELECT a.ID, a.APPOINTMENT_TYPE, a.STARTDATE, a.BRANCH_IDENTIFIER as BRANCH, p.TITLE, p.ID as PID, HASHBYTES('SHA2_256',CAST (p.ID AS VARCHAR(max))) as GENID, 
                p.GIVEN, p.SURNAME, p.EMAIL, p.NOEMAIL, p.BIRTHDATE, u.GIVEN_NAME as O_GIVEN, u.SURNAME as O_SURNAME
            FROM dbo.APPOINTMENT a with (NOLOCK), dbo.PATIENT p with (NOLOCK), dbo.USERS u with (NOLOCK)
            WHERE a.STARTDATE between '{start_date}' and '{end_date}'
            AND p.ID = a.PATIENTID 
            AND u.IDENTIFIER = a.RESID""",
        "script"    :   """
            SELECT p.TITLE, p.GIVEN, p.SURNAME, p.EMAIL, p.NOEMAIL, p.BIRTHDATE, p.BRANCH_IDENTIFIER as BRANCH
            FROM dbo.PATIENT p with (NOLOCK), dbo.CONTACT_RX c with (NOLOCK)
            WHERE c.EXPIRY_DATE='{expiry_date}'
            AND p.ID = c.PATIENTID"""   
            }

    def __init__(self):
        # SQL Connect Info
        server      = config.MSSQL_DB['server']
        database    = config.MSSQL_DB['database']
        username    = config.MSSQL_DB['username']
        password    = config.MSSQL_DB['password'] 
        cnxn        = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
        self.cursor = cnxn.cursor()

    def getFeedbackCx(self):
        """ Retrieve Cx """
        # Pre Params date always assumed to be date of running the function
        s_date = datetime.date.today()
        s_date = s_date + datetime.timedelta(days=config.DATE_OFFSET['feedback'])
        start_date = f"{s_date} 00:00:00" 
        end_date = f"{s_date} 23:59:59"

        sql = self.sqlstrings['feedback'].replace('{start_date}',start_date)
        sql = sql.replace('{end_date}',end_date)
        customers = self.filteredCustomers(self.cursor.execute(sql).fetchall())
        return customers

    def getFollowupCx(self):
        """ Retrieve Cx """
        # Pre Params date always assumed to be date of running the function
        s_date = datetime.date.today()
        s_date = s_date + datetime.timedelta(days=config.DATE_OFFSET['followup'])
        start_date = f"{s_date}" 
        print(start_date)
        sql = self.sqlstrings['followup'].replace('{collected_date}',start_date)
        customers = self.filteredCustomers(self.cursor.execute(sql).fetchall())
        return customers

    def getPEQPx(self):
        s_date = datetime.date.today()
        s_date = s_date + datetime.timedelta(days=config.DATE_OFFSET['peq'])
        start_date = f"{s_date} 00:00:00" 
        end_date = f"{s_date} 23:59:59"
        sql = self.sqlstrings['peq'].replace('{start_date}',start_date)
        sql = sql.replace('{end_date}',end_date)
        customers = self.filteredCustomers(self.cursor.execute(sql).fetchall())
        return customers


    def getScriptPx(self):
        e_date = datetime.date.today() 
        e_date = f"{e_date} 00:00:00.000"
        print(e_date)
        customers = []
        sql = self.sqlstrings['script'].replace('{expiry_date}',e_date)
        customers = self.filteredCustomers(self.cursor.execute(sql).fetchall())
        return customers

    def above_17(self, dob):
        today = datetime.date.today() 
        age = today.year - dob.year - ((today.month, today.day) <  (dob.month, dob.day)) 
        if (age > 17):
            return True
        return False

    def age(self, dob):
        today = datetime.date.today() 
        return today.year - dob.year - ((today.month, today.day) <  (dob.month, dob.day))        

    def filteredCustomers(self,customers):
        """ Basic filter to avoid sending mail to <18's and People with the NOEMAIL flag on """
        tmp = []
        for c in customers:
            if ((c.EMAIL != '') and (c.NOEMAIL == False) and (self.above_17(c.BIRTHDATE))): 
                tmp.append(c)
        return tmp