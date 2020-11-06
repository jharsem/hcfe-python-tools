#   class Optomate Benefit
#   Benefit Specific Optomate CLass
#   @author Jon Harsem <jonh@hcfeye.com.au>
#   @version 1.0
#   @date 2020-09-22

import os
import pyodbc 
import datetime
import config

class OptomateBenefit():
    """ This class handles the Benefit information to and from Optomate Touch """

    sqlstrings = {
       "name"   :   """SELECT p.title, p.given, p.surname
            FROM dbo.patient p with (NOLOCK)
            WHERE p.member_number = '{memberno}'
            """
            }

    def __init__(self):
        # SQL Connect Info
        server      = config.MSSQL_DB['server']
        database    = config.MSSQL_DB['database']
        username    = config.MSSQL_DB['username']
        password    = config.MSSQL_DB['password'] 
        cnxn        = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
        self.cursor = cnxn.cursor()

    def getNamefromMemberNo(self,memberno):
        sql = self.sqlstrings['name'].replace('{memberno}',memberno)
        name = self.cursor.execute(sql).fetchone()
        if (name):
            return name
        else:
            return None
    
