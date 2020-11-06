import os
import pyodbc 
import calendar
import datetime
import smtplib
# For guessing MIME type based on file name extension
import mimetypes
from email.message import EmailMessage
import csv
import report_config as config

dispensers = []
branches = ['BLA','BON','BRO','CHA','HUR','PAR','SYD']
reports = ['#cust','#sale','#fr','$fr','#lens','$lens','#suns','$suns','#own_fr','#own_fr_cx','#lens_ng','$mgmt_d','$match_d','$total_d','$unpaid','#r/make','#w_days','#wrty','#cncl','#px_2pr','#fit','#check','#pack','#deliv','#priced']
tmp = {}
filename ={}
dirty = {}

# Get date ranges
#year = 2020
#previous_month = datetime.date.today().month - 1 or 12
#dt_from = datetime.datetime(year,previous_month,1)
#dt_to = datetime.datetime(year,previous_month,calendar.monthrange(year,previous_month)[1])

dt_from = datetime.date(2019,7,1)
dt_to = datetime.date(2020,6,30)

str_dt_from =  dt_from.strftime("%Y-%m-%d 00:00:00")
str_dt_to = dt_to.strftime("%Y-%m-%d 23:59:59")

# Cust = CALCULATE(DISTINCTCOUNT(INVOICE[PATIENTID]),INVOICE[INVOICE_ITEMS.STOCK_TYPE] in {2,3,4},INVOICE[TYPE]=2,INVOICE[DESPENSE_TOTAL]>0)
# Sale = CALCULATE(DISTINCTCOUNT(INVOICE[ID]),INVOICE[INVOICE_ITEMS.STOCK_TYPE] in {2,3,4},INVOICE[TYPE] = 2) - CALCULATE(DISTINCTCOUNT(INVOICE[ID]),INVOICE[INVOICE_ITEMS.STOCK_TYPE] in {2,3,4},INVOICE[TYPE] = 6)
# Fr = CALCULATE(SUM(INVOICE[INVOICE_ITEMS.QTY]),INVOICE[INVOICE_ITEMS.STOCK_TYPE] in {2,3}, INVOICE[TYPE] in {2}, INVOICE[DESPENSE_TOTAL]>0, INVOICE[INVOICE_ITEMS.ITEM_BARCODE] <> "0001") 
#$ Fr = CALCULATE(SUMX(INVOICE,(INVOICE[INVOICE_ITEMS.QTY] * INVOICE[INVOICE_ITEMS.UNITPRICE]) - INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]),INVOICE[INVOICE_ITEMS.STOCK_TYPE]=2,INVOICE[TYPE]=2) + CALCULATE(SUMX(INVOICE,(INVOICE[INVOICE_ITEMS.QTY] * INVOICE[INVOICE_ITEMS.UNITPRICE]) + INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]),INVOICE[INVOICE_ITEMS.STOCK_TYPE]=2,INVOICE[TYPE]=6)
# Lens = CALCULATE(SUM(INVOICE[INVOICE_ITEMS.QTY]),INVOICE[INVOICE_ITEMS.STOCK_TYPE]=4,INVOICE[TYPE] in {2,6},LEFT(INVOICE[INVOICE_ITEMS.ITEM_BARCODE],3) <> "GAP")
#$ Lens = CALCULATE(SUMX(INVOICE, (INVOICE[INVOICE_ITEMS.QTY] * INVOICE[INVOICE_ITEMS.UNITPRICE] - INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT])),INVOICE[INVOICE_ITEMS.STOCK_TYPE]=4,INVOICE[TYPE] in {2,6}) 
# Suns = CALCULATE(SUM(INVOICE[INVOICE_ITEMS.QTY]),INVOICE[TYPE] in {2,6} ,INVOICE[INVOICE_ITEMS.STOCK_TYPE]=3,INVOICE[DISCOUNTID] <> 106)
#$ Sun = CALCULATE(SUMX(INVOICE,INVOICE[INVOICE_ITEMS.QTY]*INVOICE[INVOICE_ITEMS.UNITPRICE] - INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]),INVOICE[INVOICE_ITEMS.STOCK_TYPE] = 3 , INVOICE[TYPE] in {2,6})
# Own Fr = CALCULATE(SUM(INVOICE[INVOICE_ITEMS.QTY]),INVOICE[TYPE] in {2,6}, INVOICE[INVOICE_ITEMS.ITEM_BARCODE]="0001")
# Own Fr Cx = CALCULATE(DISTINCTCOUNT(INVOICE[PATIENTID]),INVOICE[TYPE] = 2, INVOICE[INVOICE_ITEMS.ITEM_BARCODE]="0001") - CALCULATE(DISTINCTCOUNT(INVOICE[PATIENTID]),INVOICE[TYPE] = 6, INVOICE[INVOICE_ITEMS.ITEM_BARCODE]="0001")
# Lens NG = CALCULATE(SUM(INVOICE[INVOICE_ITEMS.QTY]),INVOICE[TYPE] in {2,6}, INVOICE[INVOICE_ITEMS.STOCK_TYPE]=4, LEFT(INVOICE[INVOICE_ITEMS.ITEM_BARCODE],2)="NG")
#$ Mgmt D = CALCULATE(SUM(INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]),INVOICE[DISCOUNTID]=71, INVOICE[TYPE] = 2) - CALCULATE(SUM(INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]),INVOICE[DISCOUNTID]=71, INVOICE[TYPE] = 6)
#$ Match D = CALCULATE(SUM(INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]),INVOICE[DISCOUNTID] in {77,89}, INVOICE[TYPE] = 2) - CALCULATE(SUM(INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]),INVOICE[DISCOUNTID] in {77,89}, INVOICE[TYPE] = 6)
#$ Total D = CALCULATE(SUM(INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]), INVOICE[TYPE] = 2, INVOICE[INVOICE_ITEMS.STOCK_TYPE] in {2,3,4}) - CALCULATE(SUM(INVOICE[INVOICE_ITEMS.DISCOUNT_AMOUNT]), INVOICE[TYPE] = 6, INVOICE[INVOICE_ITEMS.STOCK_TYPE] in {2,3,4})
#$ Unpaid = CALCULATE(SUMX(INVOICE,INVOICE[STOCK_CHARGED]-INVOICE[STOCK_PAID]),INVOICE[TYPE] in {2,6})
# R/Make = CALCULATE(COUNT(SPECTACLE_ORDERS[ID]),SPECTACLE_ORDERS[WARRANTY]="P")
# W'Days = CALCULATE(DISTINCTCOUNT(CLOCKINOUT[TIMESTMP])) 
# W'ty = CALCULATE(COUNT(SPECTACLE_ORDERS[ID]),SPECTACLE_ORDERS[WARRANTY] in {"S","P"})
# Cncl = CALCULATE(COUNT(SPECTACLE_ORDERS[ID]),SPECTACLE_ORDERS[CANCELLED]=TRUE())
# Px(2Pr) = CALCULATE(DISTINCTCOUNT(INVOICE[PATIENTID]),INVOICE[DISCOUNTID] in {26,74,75,97,105}, INVOICE[TYPE] = 2)
# Fit = CALCULATE(COUNT(SPECTACLE_ORDERS[ID]),USERELATIONSHIP(USERS[IDENTIFIER],SPECTACLE_ORDERS[STAFFFIT]))
# Check = CALCULATE(COUNT(SPECTACLE_ORDERS[ID]),USERELATIONSHIP(USERS[IDENTIFIER],SPECTACLE_ORDERS[STAFFCHK]))
# Pack = CALCULATE(COUNT(SPECTACLE_ORDERS[ID]),USERELATIONSHIP(USERS[IDENTIFIER],SPECTACLE_ORDERS[STAFFPCK]))
# Deliv = CALCULATE(COUNT(SPECTACLE_ORDERS[ID]),USERELATIONSHIP(USERS[IDENTIFIER],SPECTACLE_ORDERS[STAFFDEL]))

# Report column definitions   
sqlstrings = {
    "#cust":"""SELECT i.user_added,count(distinct(i.PATIENTID)) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
        AND i.BRANCH_IDENTIFIER='{}'
        AND ii.STOCK_TYPE in (2,3,4)
        AND i.TYPE = 2
        AND i.DESPENSE_TOTAL >0
        GROUP BY i.USER_ADDED""",
    "#sale":"""SELECT one.user_added, (ISNULL(one.sales,0) - ISNULL(two.sales,0))
        FROM (
            SELECT i.USER_ADDED, COUNT(DISTINCT i.ID) as sales
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN INVOICE_ITEMS ii on ii.INVOICEID = i.ID 
            WHERE SALE_DATE BETWEEN '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            and i.TYPE = 2
            and ii.STOCK_TYPE in (2,3,4)
            GROUP BY i.user_added
        ) AS one
        LEFT JOIN 
        (
            SELECT i.USER_ADDED, COUNT(DISTINCT i.ID) as sales
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN INVOICE_ITEMS ii on ii.INVOICEID = i.ID 
            WHERE SALE_DATE BETWEEN '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            and i.TYPE = 6
            and ii.STOCK_TYPE in (2,3,4)
            GROUP BY i.user_added
            ) AS two
        ON one.USER_ADDED = two.USER_ADDED""",  
    "#fr":"""SELECT i.user_added, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND ii.STOCK_TYPE in (2,3)
            AND i.TYPE = 2
            AND i.DESPENSE_TOTAL > 0
            AND ii.ITEM_BARCODE <> '0001'
        GROUP BY i.USER_ADDED""",
    "$fr":"""SELECT i.user_added, SUM(ii.QTY * ii.UNITPRICE - ii.DISCOUNT_AMOUNT) 
	    FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND ii.STOCK_TYPE = 2
            AND i.TYPE in (2,6)
	    GROUP BY i.USER_ADDED""",
    "#lens":"""SELECT i.user_added, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND ii.STOCK_TYPE = 4
            AND i.TYPE in (2,6)
            AND LEFT(ii.ITEM_BARCODE,3) <> 'GAP'
        GROUP BY i.USER_ADDED""",
    "$lens":"""SELECT i.user_added, SUM(ii.QTY * ii.UNITPRICE - ii.DISCOUNT_AMOUNT) 
	    FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND ii.STOCK_TYPE = 4
            AND i.TYPE in (2,6)
	    GROUP BY i.USER_ADDED""",
    "#suns":"""SELECT i.user_added, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND ii.STOCK_TYPE = 3
            AND i.TYPE in (2,6)
            AND i.DISCOUNTID <> 106
        GROUP BY i.USER_ADDED""",
    "$suns":"""SELECT i.user_added, SUM(ii.QTY * ii.UNITPRICE - ii.DISCOUNT_AMOUNT) 
	    FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND ii.STOCK_TYPE = 3
            AND i.TYPE in (2,6)
	    GROUP BY i.USER_ADDED""",
    "#own_fr":"""SELECT i.user_added, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND ii.ITEM_BARCODE='0001'
            AND i.TYPE in (2,6)
        GROUP BY i.USER_ADDED""",
    "#own_fr_cx":"""SELECT one.user_added, (ISNULL(one.fr,0) - ISNULL(two.fr,0)) as fr
        FROM 
            (SELECT i.user_added, COUNT(DISTINCT(i.PATIENTID)) as fr
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
            WHERE i.SALE_DATE between '{}' and '{}'
                AND i.BRANCH_IDENTIFIER='{}'
                AND ii.ITEM_BARCODE='0001'
                AND i.TYPE = 2
            GROUP BY i.USER_ADDED) as one 
        LEFT JOIN           
            (SELECT i.user_added,COUNT(DISTINCT(i.PATIENTID)) as fr 
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
            WHERE i.SALE_DATE between '{}' and '{}'
                AND i.BRANCH_IDENTIFIER='{}'
                AND ii.ITEM_BARCODE='0001'
                AND i.TYPE = 6
            GROUP BY i.USER_ADDED) as two           
        ON one.user_added = two.user_added""",
    "#lens_ng":"""SELECT i.user_added, SUM(ii.QTY) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND LEFT(ii.ITEM_BARCODE,2)='NG'
            AND i.TYPE in (2,6)
        GROUP BY i.USER_ADDED""",
    "$mgmt_d":"""SELECT one.user_added, (ISNULL(one.disc,0) - ISNULL(two.disc,0)) as disc
        FROM 
            (Select i.user_added, SUM(ii.DISCOUNT_AMOUNT) as disc
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
            WHERE i.SALE_DATE between '{}' and '{}'
                AND i.BRANCH_IDENTIFIER='{}'
                AND i.DISCOUNTID = 71
                AND i.TYPE = 2
            GROUP BY i.USER_ADDED) as one         
        LEFT JOIN           
            (Select i.user_added,SUM(ii.DISCOUNT_AMOUNT) as disc 
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
            WHERE i.SALE_DATE between '{}' and '{}'
                AND i.BRANCH_IDENTIFIER='{}'
                AND i.DISCOUNTID = 71
                AND i.TYPE = 6
            GROUP BY i.USER_ADDED) as two     
        ON one.user_added = two.user_added""",
    "$match_d":"""SELECT one.user_added, (ISNULL(one.disc,0) - ISNULL(two.disc,0)) as disc
        FROM 
            (Select i.user_added, SUM(ii.DISCOUNT_AMOUNT) as disc
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
            WHERE i.SALE_DATE between '{}' and '{}'
                AND i.BRANCH_IDENTIFIER='{}'
                AND i.DISCOUNTID in (77,89)
                AND i.TYPE = 2
            GROUP BY i.USER_ADDED) as one         
        LEFT JOIN           
            (Select i.user_added,SUM(ii.DISCOUNT_AMOUNT) as disc 
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
            WHERE i.SALE_DATE between '{}' and '{}'
                AND i.BRANCH_IDENTIFIER='{}'
                AND i.DISCOUNTID in (77,89)
                AND i.TYPE = 6
            GROUP BY i.USER_ADDED) as two     
        ON one.user_added = two.user_added""",
    "$total_d":"""SELECT one.user_added, (ISNULL(one.disc,0) - ISNULL(two.disc,0)) as disc
        FROM 
            (Select i.user_added, SUM(ii.DISCOUNT_AMOUNT) as disc
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
            WHERE i.SALE_DATE between '{}' and '{}'
                AND i.BRANCH_IDENTIFIER='{}'
                AND ii.STOCK_TYPE in (2,3,4)
                AND i.TYPE = 2
            GROUP BY i.USER_ADDED) as one         
        LEFT JOIN           
            (Select i.user_added,SUM(ii.DISCOUNT_AMOUNT) as disc 
            FROM dbo.INVOICE i with (NOLOCK)
            JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
            WHERE i.SALE_DATE between '{}' and '{}'
                AND i.BRANCH_IDENTIFIER='{}'
                AND ii.STOCK_TYPE in (2,3,4)
                AND i.TYPE = 6
            GROUP BY i.USER_ADDED) as two     
        ON one.user_added = two.user_added""",
    "$unpaid":"""SELECT i.user_added, SUM(i.STOCK_CHARGED-i.STOCK_PAID) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND i.TYPE in (2,6)
        GROUP BY i.USER_ADDED""",
    "#r/make":"""SELECT so.user_added, COUNT(so.ID) 
        FROM dbo.SPECTACLE_ORDERS so with (NOLOCK)
        WHERE so.ORDER_DATE between '{}' and '{}'
            AND so.BRANCH_IDENTIFIER='{}'
            AND so.WARRANTY='P'
        GROUP BY so.USER_ADDED""",
    "#w_days":"""SELECT cl.user_added, COUNT(cl.TIMESTMP) 
        FROM dbo.CLOCKINOUT cl with (NOLOCK)
        WHERE cl.TIMESTMP between '{}' and '{}'
            AND cl.BRANCH_IDENTIFIER='{}'
        GROUP BY cl.USER_ADDED""",
    "#cncl":"""SELECT so.user_added, COUNT(so.ID) 
        FROM dbo.SPECTACLE_ORDERS so with (NOLOCK)
        WHERE so.ORDER_DATE between '{}' and '{}'
            AND so.BRANCH_IDENTIFIER='{}'
            AND so.CANCELLED=1
        GROUP BY so.USER_ADDED""",
    "#wrty":"""SELECT so.user_added, COUNT(so.ID) 
        FROM dbo.SPECTACLE_ORDERS so with (NOLOCK)
        WHERE so.ORDER_DATE between '{}' and '{}'
            AND so.BRANCH_IDENTIFIER='{}'
            AND so.WARRANTY in ('S','P')
        GROUP BY so.USER_ADDED""",
    "#px_2pr":"""SELECT i.user_added, COUNT(DISTINCT(i.PATIENTID)) 
        FROM dbo.INVOICE i with (NOLOCK)
        JOIN dbo.INVOICE_ITEMS ii ON ii.INVOICEID = i.ID
        WHERE i.SALE_DATE between '{}' and '{}'
            AND i.BRANCH_IDENTIFIER='{}'
            AND i.TYPE = 2
            AND i.DISCOUNTID in (26,74,75,97,105)
        GROUP BY i.USER_ADDED""",
    "#fit":"""SELECT so.STAFFFIT, COUNT(so.ID) 
        FROM dbo.SPECTACLE_ORDERS so with (NOLOCK)
        WHERE so.ORDER_DATE between '{}' and '{}'
            AND so.BRANCH_IDENTIFIER='{}'
        GROUP BY so.STAFFFIT""",                
    "#check":"""SELECT so.STAFFCHK, COUNT(so.ID) 
        FROM dbo.SPECTACLE_ORDERS so with (NOLOCK)
        WHERE so.ORDER_DATE between '{}' and '{}'
            AND so.BRANCH_IDENTIFIER='{}'
        GROUP BY so.STAFFCHK""",  
    "#pack":"""SELECT so.STAFFPCK, COUNT(so.ID) 
        FROM dbo.SPECTACLE_ORDERS so with (NOLOCK)
        WHERE so.ORDER_DATE between '{}' and '{}'
            AND so.BRANCH_IDENTIFIER='{}'
        GROUP BY so.STAFFPCK""",
    "#deliv":"""SELECT so.STAFFDEL, COUNT(so.ID) 
        FROM dbo.SPECTACLE_ORDERS so with (NOLOCK)
        WHERE so.ORDER_DATE between '{}' and '{}'
            AND so.BRANCH_IDENTIFIER='{}'
        GROUP BY so.STAFFDEL""",
    "#priced":"""SELECT sa.USER_ADDED, COUNT(sa.ID) 
        FROM dbo.STOCK_ARRIVAL sa with (NOLOCK)
        WHERE sa.ARRIVAL_DATE between '{}' and '{}'
            AND sa.BRANCH_IDENTIFIER='{}'
        GROUP BY sa.USER_ADDED"""        
    }

def findValueForDispenser(disp,data):
    for (b,d) in data:
        if (b == disp):
            dirty[disp] = 1
            return d
    return 0

# Get all Dispensers
def getDispensers():
    sql = """SELECT distinct(u.identifier) 
    FROM dbo.USERS u with (NOLOCK)
        WHERE u.user_type in (2,4,5) 
        AND u.inactive=0
        AND u.identifier not in ('JH','KD','LVA','TLO')"""
    return cursor.execute(sql).fetchall() 

# Build CSV
def buildCSV(fr,to,title,branch):
    filename = "report-"+title+"-"+fr+"-to-"+to+"-"+branch+".csv"
    with open(filename, "w", newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        header = ['Dispenser Code']
        for r in reports:
            header.append(r)
        csv_writer.writerow(header) # write headers
        for dispenser in dispensers:
            if (dispenserHasValues(dispenser[0])):
                row = [dispenser[0]]
                for r in reports:
                    row.append(findValueForDispenser(dispenser[0],tmp[r]))
                csv_writer.writerow(row)
    return filename

def dispenserHasValues(d):
    for r in reports:
        if (findValueForDispenser(d,tmp[r]) != 0):
            return 1
    return 0

# Send Email
def sendReport(f):
    fromaddr = "Report Robot <reports@hcfeye.com.au>"
    #toaddrs  = "Jon Harsem <jonh@hcfeye.com.au>"
    toaddrs  = "Lauren Valentine <laurenv@hcfeye.com.au>"

    msg = EmailMessage()
    messageText = """
    Your report is based on the date range from {} to {}
    
    Report Notes
     
        #cust       Distinct speclens customers (t.i.a. sun)
        #sale       distinct Speclens invoices (t.i.a. returns)
        #fr         Frame line item qty less own frame (inc sun, - return)         
        $fr         Frame $ (t.i.a. discount, - sun!, t.i.a. return)
        #lens       Lens line item qty (t.i.a. return, t.i.a. GAP items)
        $lens       Lens $ (t.i.a. discount, t.i.a. return)
        #suns       Suns line item qty (t.i.a. discount 106, t.i.a. returns)
        $suns       Suns $ (t.i.a.discount, t.i.a.returns)
        #own_fr     Own frame (t.i.a. returns)
        #own_fr_cx  Unique Px sales with own frame (t.i.a. returns)
        #lens_ng    Sales of NG* barcode (t.i.a. returns)
        $mgmt_d     Managers Discount (t.i.a. returns, Type 71)
        $match_d    Price Match Discount (t.i.a. returns, Type 77,89)
        $total_d    Total discount given on speclens items (t.i.a. returns)
        $unpaid     Unpaid $ 
        #r/make     Total specjobs with warranty marked 'P'
        #w_days     Total days worked
        #wrty       Total Warranty (Supplier or Problem) 
        #cncl       Total specjobs marked cancelled
        #px_2pr     Total distinct px with invoice discount in (26,74,75,97,105)
        #fit        Total specjobs fitted by dispenser (fit field)
        #check      Total specjobs checked by dispenser (check field)
        #pack       Total specjobs packed by dispenser (packed field)
        #deliv      Total specjobs delivered by dispenser (delivered field)
        #priced     Total stock arrivals by dispenser
    
    """.format(str_dt_from, str_dt_to)
    msg.set_content(messageText)
    msg['Subject'] = "HCFE Dispenser Report for {} to {}".format(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"))
    msg['From'] = fromaddr
    msg['To'] = toaddrs
    # Send the message via our own SMTP server.
    s = smtplib.SMTP('192.168.2.41')
    for b in branches:
        path = os.path.join('.', f[b])
        ctype, encoding = mimetypes.guess_type(path)
        if ctype is None or encoding is not None:
            # No guess could be made, or the file is encoded (compressed), so
            # use a generic bag-of-bits type.
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        with open(path, 'rb') as fp:
            msg.add_attachment(fp.read(), maintype=maintype, subtype=subtype, filename=f[b])
    s.send_message(msg)
    s.quit()
    return

# Global

# SQL Connect Info
server = config.MSSQL_DB['server']
database = config.MSSQL_DB['database']
username = config.MSSQL_DB['username']
password = config.MSSQL_DB['password'] 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Build SQL string for Purchase customers
print ('Timeframe',str_dt_from,str_dt_to)
dispensers = getDispensers()
for b in branches:
    print("Report for {} ...".format(b))
    for r in reports:
        if (sqlstrings[r].count('{}') == 3):
            sql = sqlstrings[r].format(str_dt_from,str_dt_to,b)
        if (sqlstrings[r].count('{}') == 6):
            sql = sqlstrings[r].format(str_dt_from,str_dt_to,b,str_dt_from,str_dt_to,b) 
        result = cursor.execute(sql).fetchall()
        tmp[r] = result
    filename[b] = buildCSV(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"),'dispensers',b)
    print("done!")
# Process Rows.
print ("------------")
sendReport(filename)
for b in branches:
    os.rename(filename[b], './data/'+filename[b])
