url='http://test-iis-eyecare.hcf.local/Eyecare.Main/default?=00999~SEARCH~~{}~{}~D~{}~{}'
surname='SMITH'
initial='J'
dob='' #DDMMYYYY
pcode='~~~~'

def pad(type,str):
    if type=='s':
        return '{0:~<20}'.format(str)
    if type=='i':
        return '{0:~<2}'.format(str)
    if type=='d':
        return '{0:~<8}'.format(str)
    return 1        
print('012345678901234567890123')
print(pad('s',surname))
print(pad('i',initial))
print(pad('d',dob))
print('---------------------')
print(url.format(pad('s',surname),pad('i',initial),pad('d',dob),pcode))
