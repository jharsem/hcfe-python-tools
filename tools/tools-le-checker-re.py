# /^Return\s{11}:\s[0-9]{30}([0-9]+)(DEP~~([0-9]{8})~([0-9]{8})~([A-Z]{1})~([A-Z\s]{24})~[A-Z\s]{20}~([A-Z]{1})~([0-9]{8})~([1-13]{1})~([0-9]{8})~([0-9\s]{8})~([$\.0-9\s]{8})~([$\.0-9\s]{8})~([A-Z]{1})~([$\.0-9\s]{8}))/g
import sys
import re 
import requests
import config 
from requests.exceptions import HTTPError
from pprint import pprint

def parseStr(query_type,str):
    exp = config.query[query_type]['re']
    fields = config.query[query_type]['fields']
    arr = re.findall(exp,str)
    for a in arr:
        i = 0
        print(a)
        for field in fields:
            print(field,a[i])
            i=i+1

def main():    
    if (len(sys.argv) != 3):
        print ("""
        Limit Enquiry NH Checker
        by Jon Harsem <jonh@hcfeye.com.au>

        This tool requires a query type and query parameter
        ie. tools-le-checker.py DEP 12345678""")
        exit(0)
    query_type = sys.argv[1].lower()
    parameter = sys.argv[2]
    r = {}
    test_type = 'test'
    url = config.base_urls[test_type] + config.query[query_type]['url'].format(parameter)
    # sending get request and saving the response as response object 
    print(url)
    try:
        r[test_type] = requests.get(url)
        # If the response was successful, no Exception will be raised
        r[test_type].raise_for_status()
    except HTTPError as http_err:
        print(f'HTTP error occurred: {http_err}')  # Python 3.6
    except Exception as err:
        print(f'Other error occurred: {err}')  # Python 3.6
    else:
        print(r[test_type].text)
        print('--------')
        parseStr(query_type,r[test_type].text)
     
    # Start your Comparison Engines!
   
    filename = './data/'+test_type+'_'+query_type+"-"+parameter+'.txt'
    outfile = open( filename, 'w' )
    outfile.write(r[test_type].text)
    outfile.close()
if __name__ == '__main__':
    main()
