#
# Config file for python3 Benefit tools
# DO NOT SHARE!
#
CONN = {
    "user":"Eyecare",
#   TEST    
#    "passwd":"3y3c@r3toHCF!",
#    "host":"mft4.hcfbiz.com.au",
#   Local TEST
#    "host":"localhost",
#    "port":22,
#
#   PROD
    "passwd":"s84n8#TNpnpp",
    "host":"mft3.hcfbiz.com.au",
    "port":9922
}

MSSQL_DB = {
    'server':'tcp:dbserver.hcfeye.com.au', 
    'database':'Optomate', 
    'username':'sa', 
    'password':'Y2xDY7wVm-@&x7yL' 
}

EMAIL = {
    'from':'Report Robot <reports@hcfeye.com.au>',
    'host':'192.168.2.41'
}

BRANCHES = {
    "BLA":{
        "prov_code":"A132380",
        "report_contact":"blacktown@hcfeye.com.au"
    },
    "BON":{
        "prov_code":"A118508",
        "report_contact":"bondijunction@hcfeye.com.au"
    },
    "BRO":{
        "prov_code":"A128129",
        "report_contact":"brookvale@hcfeye.com.au"
    },
    "CHA":{
        "prov_code":"A112609",
        "report_contact":"chatswood@hcfeye.com.au"
    },
    "HUR":{
        "prov_code":"A106698",
        "report_contact":"hurstville@hcfeye.com.au"
    },
    "PAR":{
        "prov_code":"A105322",
        "report_contact":"parramatta@hcfeye.com.au"
    },
    "SYD":{
        "prov_code":"A100832",
        "report_contact":"reception@hcfeye.com.au"
    }
}

PATTERNS = {
    # general reports filename pattern
    'benefitreports':'(\\d{12})-(A\\d{6})-H20017-(\\w*)\\.txt',
    # summary report claim detail pattern
    'sl_pattern':'(\\d+)\\s*(\\w)\\s*(\\d+)\\s+(\\d+)\\s(\\w+)\\s+([\\w//]+)\\s+(\\d+)\\s+(.*)\\n',
    # summary report claim detail headers
    's_header':['Name','MemberNo','Suffix','ItemCode','Policy','H/Policy','Date Service','Benefit Paid','Message'],
    # summary report subheaders pattern
    'sh_pattern':'(^\\w+\\s*\\w*)\\s:$',
    # summary report subheader detail pattern
    'shd_pattern':'^\\t{2}(\\w+)\\s:\\s(.+)',
    # benefit line 
    'benefit_pattern':'((\\d{8})(\\d{8})(\\d{3})\\s{2}(\\d{6})(\\d{6})(\\d{6})\\s(A\d{6})\\s{7})'
}