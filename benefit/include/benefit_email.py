import config
import smtplib
# For guessing MIME type based on file name extension
import mimetypes
from email.message import EmailMessage
from include import benefit_config as config
import os

def sendBenefits(files,to,subject,body,pickup_date,branch):
    fromaddr = config.EMAIL['from']
    #toaddrs  = "laurenv@hcfeye.com.au,krisd@hcfeye.com.au"
    if (to==''):
        to  = "jonh@hcfeye.com.au"
    msg = EmailMessage()
    messageText = body.format(pickup_date, branch)
    msg.set_content(messageText)
    msg['Subject'] = subject.format(pickup_date)
    msg['From'] = fromaddr
    msg['To'] = to
    # Send the message via our own SMTP server.
    s = smtplib.SMTP(config.EMAIL['host'])
    for f in files:
        print("file to attach",f)
        path = os.path.join('.', f)
        ctype, encoding = mimetypes.guess_type(path)
        if ctype is None or encoding is not None:
            # No guess could be made, or the file is encoded (compressed), so
            # use a generic bag-of-bits type.
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        with open(path, 'rb') as fp:
            msg.add_attachment(fp.read(), maintype=maintype, subtype=subtype, filename=f)
    s.send_message(msg)
    s.quit()
    return