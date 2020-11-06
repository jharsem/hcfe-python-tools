import sys
import difflib
import requests 
from requests.exceptions import HTTPError
from pprint import pprint

#base_test = 'http://prod-iis-eyecare.hcf.local/Eyecare.Main/default?=00999~'
#base_prod = 'http://test-iis-eyecare.hcf.local/Eyecare.Main/default?=00999~'
base_test = 'http://intranet.hcfeye.com.au/jon/test1.txt?='
base_prod = 'http://intranet.hcfeye.com.au/jon/test.txt?='

urls={
    'dep':'DEP~~{}',
    'gaps':'GAPS~~{}',
    'histc':'HISTC~~{}',
    'histp':'HISTP~~{}',
    'details':'DETAILS~~{}',
    'search':'SEARCH~~{}~{}~D~{}~{}'
}
'''
00999~DEP~~policy number
00999~DETAILS~~policy number
00999~HISTC~~policy number
00999~HISTP~~member number
00999~GAPS~~member number
Session Number is always 00999
Search format is 0099~SEARCH~~surname~initial~D~date of birth~postcode
'''

def main():    
    if (len(sys.argv) != 3):
        print ("""
        Limit Enquiry NH Checker
        by Jon Harsem <jonh@hcfeye.com.au>

        This tool requires a query type and query parameter
        ie. tools-le-checker.py DEP 12345678""")
        exit(0)
    query_type = sys.argv[1]
    parameter = sys.argv[2]
    r = {}
    url = base_prod + urls[query_type.lower()].format(parameter)
    # sending get request and saving the response as response object 
    print(url)
    try:
        r['prod'] = requests.get(url)

        # If the response was successful, no Exception will be raised
        r['prod'].raise_for_status()
    except HTTPError as http_err:
        print(f'HTTP error occurred: {http_err}')  # Python 3.6
    except Exception as err:
        print(f'Other error occurred: {err}')  # Python 3.6
    else:
        url = base_test + urls[query_type.lower()].format(parameter)
        # sending get request and saving the response as response object 
        print(url)
        try:
            r['test'] = requests.get(url)

            # If the response was successful, no Exception will be raised
            r['test'].raise_for_status()
        except HTTPError as http_err:
            print(f'HTTP error occurred: {http_err}')  # Python 3.6
        except Exception as err:
            print(f'Other error occurred: {err}')  # Python 3.6
        else:
            print('Success!')
    # Start your Comparison Engines!
    d = difflib.Differ()

    #result = difflib.ndiff([r['prod'].text],[r['test'].text])
    differ = difflib.HtmlDiff( tabsize=4, wrapcolumn=40 )
    html = differ.make_file( [r['prod'].text],[r['test'].text],'PROD','NH', context=False )
    filename = 'difftest_'+query_type+"-"+parameter+'.html'
    outfile = open( filename, 'w' )
    outfile.write(html)
    outfile.close() 
if __name__ == '__main__':
    main()