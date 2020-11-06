base_urls = {
    'test':'http://test-iis-eyecare.hcf.local/Eyecare.Main/default?=00999~',
    'prod':'http://prod-iis-eyecare.hcf.local/Eyecare.Main/default?=00999~'
}

query = {
    'dep':{ 'url':'DEP~~{}',
            'fields':['policy no','member no','order','first','last','gender','dob','status','start dt','term dt','a limit','o limit','N?','l boost'],
            're':'DEP~~([0-9]{8})~([0-9]{8})~([A-Z]{1})~([A-Z\s]{24})~([A-Z\s]{20})~([A-Z]{1})~([0-9]{8})~([1-9A-Z]{1})~([0-9]{8})~([0-9\s]{8})~([$\.0-9\s]{8})~([$\.0-9\s]{8})~([A-Z]{1})~([$\.0-9\s]{8})'
            },
    'gaps':{  'url':'GAPS~~{}',
              'fields':['member no','unknown','space','type','date fr','date to','reason'],
              're':'GAPS~~([0-9]{8})~([0-9]{8})~([A-Z0-9\s]{1})~([A-Z0-9]{1})~([0-9]{8})~([0-9]{8})~([A-Z0-9]{13})'},
    'histc':{ 'url':'HISTC~~{}',
              'fields':['policy no','member no','DATA?','line no?','claim no','processed dt','location','provider','A?','item code','service dt','cost','benefit'],
              're':'HISTC~~([0-9]{8})~([0-9]{8})~([A-D]DATA)~~([0-9]{5})~([0-9]{5})~([0-9]{8})~([A-Z][0-9]{5})~([A-Z0-9\s]{8})~([A-Z\s]{5})~([A-Z][0-9\s]{4})~([0-9]{8})~([\s$0-9\.]{8})~([\s$0-9\.]{8})'},
    'histp':{ 'url':'HISTP~~{}',
              'fields':[],
              're':''},
    'details':{ 'url':'DETAILS~~{}',
                'fields':['client','customer','suffix','title code','first','last','gender','dob','home ph','work-ph','h-addr1','h-addr2','h-suburb','h-pcode','h-mail-ret-date','p-addr1','p-addr2','p-suburb','p-pcode','p-mail-ret-date','status','status-date','eff-date','scale','pay-method','pay-freq','enrolment-code','waver-code','id-ind','dte-on','dte-payed-to','state','card-no','card-date-issued','p1','p2','p3','p4','p5','p6','group-dpt','group-no','sect-no','user-resp'],
                're':'DETAILS~~([0-9]{8})~([0-9]{8})~([A-Z]{1})NAME~~([\s0-9]{2})~([A-Z\s]{24})~([A-Z\s]{20})~([A-Z\s]{1})~([0-9]{8})PHONE~~([0-9\s]{15})~([0-9\s]{15})HOME~~([0-9A-Z\s\-]{26})~([0-9A-Z\s\-]{26})~([0-9A-Z\s\-]{26})~([0-9\s]{6})~([0-9A-Z\s]{8})POST~~([0-9A-Z\s\-]{26})~([0-9A-Z\s\-]{26})~([0-9A-Z\s\-]{26})~([0-9\s]{6})~([0-9A-Z\s]{8})POLICY~~([A-Z\s]{1})~([0-9]{8})~([0-9]{8})~([A-Z\s]{1})~([A-Z\s]{1})~([A-Z\s]{1})~([0-9]{2})~([0-9]{2})~([A-Z\s]{1})~([0-9]{8})~([0-9]{8})~([A-Z\s]{1})~([0-9]{2})~([0-9]{8})PRODS~~([0-9A-Z\s]{3})~([0-9A-Z\s]{3})~([0-9A-Z\s]{3})~([0-9A-Z\s]{3})~([0-9A-Z\s]{3})~([A-Z\s]{8})~~([0-9]{8})~([0-9]{5})~([0-9\s]{3})~([0-9]{3})'},
    'search':{ 'url':'SEARCH~~{}~{}~D~{}~{}',
               'fields':[],
               're':''}
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