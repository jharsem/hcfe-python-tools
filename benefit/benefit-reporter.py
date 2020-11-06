import pysftp
import os
import zipfile
import datetime
import re
import sys
import logging
from include import benefit_email as email
from include import benefit_config as config
from include.optomate_benefit import OptomateBenefit as optomate
from include.report_writer import report

'''
This python3 script grabs the exported report files from the HCF SFTP server 
Jon Harsem <jharsem@hcf.com.au> on 2020-08-31

Requirements: pysftp

If present we're looking for 3 files of the format
DDMMYYYYHHMM-BRANCH_CODE-H20017-[Error|ExceptionReport|SummaryReport].txt

eg
230620200944-A128129-H20017-Error.txt
230620200944-A128129-H20017-ExceptionReport.txt
230620200944-A128129-H20017-SummaryReport.txt

for each branch. 
'''
# Basic config

cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

today =  datetime.date.today().strftime("%d-%m-%Y")
errors = 0
files = []
cleanup = []
logging.basicConfig(filename='./benefit-reporter.log', level=logging.DEBUG, format='%(asctime)s %(levelname)s %(message)s')

# Helper Functions 

def getFiles(sftp, files):
    for f in files:
        sftp.get(f) 

def archiveLocalFiles(files):
    archive = zipfile.ZipFile('./archive/'+today+'-benefitreports.zip', 'w')
    
    for file in files:
        archive.write(file, compress_type = zipfile.ZIP_DEFLATED)
        os.remove(file)

    archive.close()

def filesForBranch(files,branch):
    tmp = []
    for f in files:
        if config.BRANCHES[branch]['prov_code'] in f:
            tmp.append(f)
    return tmp
#
# Get Reports
#
def getReports(cleanup):
    with pysftp.Connection(host=config.CONN['host'], port=config.CONN['port'], username=config.CONN['user'], 
    password=config.CONN['passwd'], log="./benefit_mover.log", cnopts=cnopts) as sftp:
        with sftp.cd('eyecare/outgoing'):
            files = sftp.listdir()
            #Get Files
            getFiles(sftp,files)
    #Instantiate Report Writer
    rWriter = report(files)
    #Generate Summary Excel file
    rWriter.build()

    #Configure Emails 
    body = "Benefit Summary Reports Picked up on {} for {} branch(es)"
    subject = "Benefit Summary Reports picked up on {} (Use this one)"
    # Find & Process the summary report
    
    # Email Summary and all files to Lauren 
    email.sendBenefits([rWriter.fileName],"jonh@hcfeye.com.au",subject,body,today,'All')
    cleanup.append(rWriter.fileName)

    # Email Summary and branch Specific files to each branch
    for f in rWriter.summaryFiles:
        filesToAttach = filesForBranch(files,f['branch'])
        filesToAttach.append(rWriter.buildOne(f['branch']))
        email.sendBenefits(filesToAttach,"jonh@hcfeye.com.au,"+config.BRANCHES[f['branch']]['report_contact'],subject,body,today,f['branch'])
        cleanup = cleanup + filesToAttach
    
    archiveLocalFiles(cleanup)

try:
    getReports(cleanup)
except:
    cleanup = []
    errors = errors + 1
    logging.warning('Try #%s',errors)
    getReports(cleanup)
    if (errors == 10):
        logging.error('Too many retries - giving up')
        sys.exit(1)

#print(all_reports)