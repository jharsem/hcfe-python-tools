#   class Template
#   @author Jon Harsem <jonh@hcfeye.com.au>
#   @version 1.0

import os
import re
import email_config as config

class Template:
    """ Template processing class """

    template_path = "./tmpl"
    template = ""
    ttype = ""

    def __init__(self, ttype):
        filename = f"{self.template_path}/{ttype}.html"
        with open(filename,'r') as f:
            self.ttype = ttype
            self.template = f.read()

    def process(self, data):
        """ Process a template and return a filled out version as a string """
        fields_obj = self.fields()     
        branch = config.BRANCHES[data.BRANCH]
        manager = branch['manager']

        for f in fields_obj['db']:
            self.template = self.template.replace("{{D_"+f+"}}",str(getattr(data,f)).title())
        
        for f in fields_obj['config']:
            if (f == "MANAGER_NAME"):
                self.template = self.template.replace("{{B_"+f+"}}",manager['name'])
                continue
            if (f == "MANAGER_EMAIL"):
                self.template = self.template.replace("{{B_"+f+"}}",manager['email'])
                continue
        
            self.template = self.template.replace("{{B_"+f+"}}",branch[str(f).lower()])
        
        if (self.ttype == "peq"):
            try:
                app_extra = config.APPOINTMENT_INFO[getattr(data,'APPOINTMENT_TYPE')]
            except KeyError:
                app_extra = ''            
            self.template = self.template.replace("{{A_APP_EXTRA}}",app_extra)


    def fields(self):
        """ introspect template and return array of field names """
        tmp = dict()
        tmp['db'] = re.findall(r'{{D_(\w+)}}',self.template)
        tmp['config'] = re.findall(r'{{B_(\w+)}}',self.template)
        return tmp