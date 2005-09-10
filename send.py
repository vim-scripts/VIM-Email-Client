############################################################
## Vim Email Client v0.2 - SMTP Sender                    ##
##                                                        ##
##                                                        ##
## Created by David Elentok.                              ##
##                                                        ##
## This software is released under the GPL License.       ##
##                                                        ##
############################################################
# -*- coding: cp1255 -*-
__version__ = "0.2"
__author__ = "David Elentok (3david@gmail.com)"
__copyright__ = "(C) 2005 David Elentok. GNU GPL 2."


############################################################
## Imports
import smtplib, sys, os.path, time, getpass
import mimetypes

from email import Encoders
from email.Message import Message
from email.MIMEAudio import MIMEAudio
from email.MIMEBase import MIMEBase
from email.MIMEMultipart import MIMEMultipart
from email.MIMEImage import MIMEImage
from email.MIMEText import MIMEText

############################################################
## Configuration
host = 'smtp.server.com'
port = 25
user = 'myusername'
passwd = getpass.getpass()
mailboxdir = "D:\\data\\mailbox"


# the mailbox that contains that sent emails
sentmbx = "%s\\sent.mbx" % mailboxdir


############################################################
def add_to_mailbox(mbox, mailFrom, mailSubject, text):
  if os.path.exists (mbox):
    fd = open(mbox)
    contents = fd.read()
    fd.close()
  else:
    contents = ""

  from_column_width = 15

  # Create Time String
  t = time.localtime()
  strtime = time.strftime("%d/%m/%Y %H:%M",t)


  # Create From String
  strfrom = mailFrom

  if len(strfrom) > from_column_width:
    strfrom=strfrom[0:from_column_width]

  strfrom = strfrom.center(from_column_width)

  # Create Subject String
  title = "* %s | %s | %s\n" % (strtime, strfrom, mailSubject)

  fd = open(mbox, "wt")


  fd.write (''.center(70,'#'));
  fd.write ('\n')
  fd.write (title);
  fd.write (text)

  if not contents == "":
    fd.write (contents)

  fd.close()

############################################################
def file2msg (file):
  ## TODO: check if file exists
  ctype, encoding = mimetypes.guess_type(file)

  if ctype is None or encoding is not None:
    ctype = 'application/octet-stream'

  maintype, subtype = ctype.split('/')

  print "==> Adding file [%s] using [%s]" % (file, ctype)

  if maintype == "text":
    fp = open(file)
    msg = MIMEText(fp.read(), _subtype = subtype)
    fp.close()
  elif maintype == "image":
    fp = open(file, 'rb')
    msg = MIMEImage(fp.read(), _subtype = subtype)
    fp.close()
  elif maintype == "audio":
    fp = open(file, 'rb')
    msg = MIMEAudio(fp.read(), _subtype = subtype)
    fp.close()
  else:
    fp = open(file, 'rb')
    msg = MIMEBase(maintype, subtype)
    msg.set_payload(fp.read())
    fp.close()
    Encoders.encode_base64(msg)

  return msg

############################################################
if len(sys.argv) < 2:
  print
  print "SMTP Sender"
  print
  print "  Correct syntax:"
  print "   %s [message_file]" % sys.argv[0]
  print
  sys.exit(1)
  
############################################################
filename = ' '.join(sys.argv[1:])
file = open(filename, 'rt')

mymsg = MIMEMultipart()

text=""
handle_attachments = False

for line in file.readlines():
  ## handle "To:" field ====================================
  if line[0:3] == "To:":
    fields = line.strip().split(':')
    mymsg['To'] = ':'.join(fields[1:])

  ## handle "From:" field ==================================
  elif line[0:5] == "From:":
    fields = line.strip().split(':')
    mymsg['From'] = ':'.join(fields[1:])

  ## handle "Subject:" field ===============================
  elif line[0:8] == "Subject:":
    fields = line.strip().split(':')
    mymsg['Subject'] = ':'.join(fields[1:])

  ## handle "**** Attachements: " ==========================
  elif line[0:17] == "#### Attachments:":
    print "######### Found attachments ###############"
    handle_attachments = True
    ## add text message to the message object:
    msg = MIMEText(text, 'plain')
    mymsg.attach(msg)

  ## handle text message contents ==========================
  elif not handle_attachments:
    text += line

  ## handle attachements ===================================
  elif line.strip() != "":
    fname = line.strip()
    msg = file2msg (fname)
    msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(fname))
    mymsg.attach(msg)

file.close()

## incase there weren't any attachments
if not handle_attachments:
  msg = MIMEText(text, 'plain')
  mymsg.attach(msg)
    
## to guarantee the message ends with a new line:
mymsg.epilogue = ''
    
sender = smtplib.SMTP(host, port)

sender.starttls()
sender.login(user, passwd)

sender.set_debuglevel(1)
sender.sendmail(mymsg['From'], mymsg['To'].split(','), mymsg.as_string())
sender.quit()

#Add the sent email to sent.mbx
file = open(filename)
add_to_mailbox(sentmbx, mymsg['To'], mymsg['Subject'], file.read())
file.close()

  

