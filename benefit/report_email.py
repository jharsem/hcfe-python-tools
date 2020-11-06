import config.py
import smtplib
# For guessing MIME type based on file name extension
import mimetypes
from email.message import EmailMessage
import os

def sendReport(f,to,subject,descr,dt_to,dt_from):
    fromaddr = config.EMAIL['from']
    #toaddrs  = "laurenv@hcfeye.com.au,krisd@hcfeye.com.au"
    if (to==''):
        to  = "jonh@hcfeye.com.au"
    msg = EmailMessage()
    messageText = descr.format(dt_from.strftime("%Y-%m-%d 00:00:00"), dt_to.strftime("%Y-%m-%d 23:59:59"))
    msg.set_content(messageText)
    msg['Subject'] = subject.format(dt_from.strftime("%Y-%m-%d"),dt_to.strftime("%Y-%m-%d"))
    msg['From'] = fromaddr
    msg['To'] = to
    # Send the message via our own SMTP server.
    s = smtplib.SMTP(config.EMAIL['host'])
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