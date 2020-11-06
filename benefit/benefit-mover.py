import pysftp
import logging
from include import benefit_config as config
import os
import glob
import datetime
import re
import sys

'''
This python3 script pushes the exported benefit files to the HCF sftp server 
Jon Harsem <jhgarsem@hcf.com.au> 2020-06-21

Requirements: pysftp

7 files of the format
19022020-A100832-H20017.txt    [Date of Service (DDMMYYYY) - Provider Number - H20017 (location fixed)]

'''
cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

today =  datetime.date.today().strftime("%d%m%Y")
error = 0
logging.basicConfig(filename='./benefit-mover.log', level=logging.DEBUG, format='%(asctime)s %(levelname)s %(message)s')

def insertSortedClaim(claims, member,claim):
    if (member in claims.keys()):
        mClaims = claims[member]
        mClaims.append(claim)
        mClaims = sorted(mClaims,key=lambda claim: claim[0],reverse=True)
        claims[member] = mClaims
    else:
        claims[member] = [claim]
    return claims

def writeBatch(claims, file):
    with open(file,'w') as f:
        for k in claims:
            for c in claims[k]:
                f.write(c[1])        

def preprocess(file):
    benefit = re.compile(config.PATTERNS["benefit_pattern"])
    with open(file) as f:
        claims = {}
        for line in f:
            results = benefit.findall(line)
            for r in results:
                claims = insertSortedClaim(claims, r[1],[r[5],r[0]])
    os.rename(file,file.replace('.txt','.raw'))
    writeBatch(claims, file)

def cleanup(files):
    for f in files:
        os.remove(f)

def upload():
    with pysftp.Connection(host=config.CONN['host'], port=config.CONN['port'], username=config.CONN['user'], 
    password=config.CONN['passwd'], log="./benefit_mover_ftp.log", cnopts=cnopts) as sftp:
        with sftp.cd('eyecare/incoming'):
            print('Remote Files',sftp.listdir())
            bfiles = glob.glob('./data/incoming/*.txt')
            print(bfiles)
            for f in bfiles:
                #preprocess(f)
                sftp.put(f)
                logging.info('Uploaded %s',f)

try:
    upload()
except:
    error = error + 1
    logging.warning('Retrying #%s', error)
    if (error > 9):
        logging.error('Too many retries - giving up')
        sys.exit(1)
    upload()
# End of program
