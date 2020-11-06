from lib.optomate import Optomate
from lib.template import Template

ot = Optomate()
customers = ot.getScriptPx()
c = customers[0]
tpl = Template('script')
tpl.process(c)    
print(tpl.template)