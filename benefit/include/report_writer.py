import datetime
import re
import os
import xlsxwriter
from include import optomate_benefit
from include import benefit_config as config
from xlsxwriter.utility import xl_rowcol_to_cell

class report:
    """
        Reporting class for benefit pickup.

        2020-09-22: Initial version. Converts the summary report for each centre into a spreadsheet.
    
    """
    wb = ''
    files = ''
    today = ''
    ot = ''
    fileName = ''
    
    # File pointers 
    summaryFiles = []
    exceptionFiles = []

    # Constructor

    def __init__(self,files):
        self.today = datetime.date.today().strftime("%Y%m%d")
        self.fileName = './data/'+self.today+'-benefitsummary.xlsx'
        
        # Create a workbook and add a worksheet.
        self.wb = xlsxwriter.Workbook(self.fileName)
        self.files = files
        self.ot = optomate_benefit.OptomateBenefit()
        self.summaryFiles = self.getSummaryFiles()
        self.exceptionFiles = self.getExceptionFiles()

    #
    # Helper Functions 
    #

    def padded(self,num):
        return num.rjust(8,'0')

    def getSummaryFiles(self):
        prog = re.compile(config.PATTERNS["benefitreports"])
        tmp = []
        for f in self.files:
            result = prog.match(f)
            if (result[3] == 'SummaryReport'):
                tmp.append({'branch':self.getBranchNamefromLocation(result[2]),'file':f})
        return tmp

    def getExceptionFiles(self):
        prog = re.compile(config.PATTERNS["benefitreports"])
        tmp = []
        for f in self.files:
            result = prog.match(f)
            if (result[3] == 'ExceptionReport'):
                tmp.append({'branch':self.getBranchNamefromLocation(result[2]),'file':f})
        return tmp

    def getBranchNamefromLocation(self,code):
        for k in config.BRANCHES.keys():
            if (config.BRANCHES[k]['prov_code'] == code):
                return k
        return ''

    def filetoWS(self,ws,file,branch):
        re_line = re.compile(config.PATTERNS['sl_pattern'])
        re_sh_line = re.compile(config.PATTERNS['sh_pattern'])
        re_shd_line = re.compile(config.PATTERNS['shd_pattern'])
        
        with open(file,'r') as f:
            next(f)
            l = 4
            print('Opening ',file)
            for line in f:
                result = re_line.match(line)                                # Is it a summary claim result
                sh_result = re_sh_line.match(line)                          # Is it a summary sub claim header
                shd_result = re_shd_line.match(line)                        # Is it a summary sub claim details 

                if (result):
                    name = self.ot.getNamefromMemberNo(self.padded(result[1]))
                    if (name):
                        ws.write(l, 0,' '.join([name[0],name[1],name[2]]))   # Write Name 
                    for i in range(1,8):
                        ws.write(l,i,result[i])                         # Write rest
                elif (sh_result):
                        ws.write(l,0,sh_result[1])
                elif (shd_result):
                        ws.write(l,1,shd_result[1])
                        ws.write(l,2,shd_result[2])
                l = l + 1        
   
    def createSheetandHeader(self,wb,branch):
        fmt_bold = wb.add_format({'bold': True})
        fmt_title= wb.add_format({'bold': True,'font_size':'14'})
        ws = wb.add_worksheet(branch)
        ws.write(0,0,'Summary Report for '+branch,fmt_title)
        ws.write(1,0,'Created on '+ self.today)
        ws.write_row(3,0,config.PATTERNS['s_header'],fmt_bold)
        return ws

    #
    #   Main Process Files 
    # 

    def build(self):
        for f in self.summaryFiles:
            ws = self.createSheetandHeader(self.wb,f['branch'])
            self.filetoWS(ws,f['file'],f['branch'])
        self.wb.close()
        return self.wb

    def buildOne(self,branch):
        for f in self.summaryFiles:
            if f['branch'] == branch:
                fN = './data/'+self.today+'-'+branch+'-benefitsummary.xlsx'
                # Create a workbook and add a worksheet.
                bwb = xlsxwriter.Workbook(fN)
                bws = self.createSheetandHeader(bwb,f['branch'])
                self.filetoWS(bws,f['file'],f['branch'])            
                bwb.close()
                return fN
        return None

    def archive(self):
        return
