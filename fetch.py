############################################################
## Vim Email Client v0.2 - POP3 Client                    ##
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
import getpass, poplib, email, os, os.path, time, sys
import email.Header
import sgmllib, cStringIO


############################################################
## Configuration
mailboxdir = 'd:\\data\\mailbox'

### NOTE: the password field is ignored at the moment
accounts = [ ('account1', 'pop3.server1.com', 'username1', 'password1'),
             ('account2', 'pop3.server2.com', 'username2', 'password2') ]

# the directory where the html2text.py file can be found
html2text_dir = mailboxdir

sys.path.append(html2text_dir)
import html2text

############################################################
def get_mailbox_uidls(mailbox):
  """This function returns a list of all of the uidl (unique id's)
     inside an .mbx file """

  mbxfile = "%s\\%s.mbx" % (mailboxdir, mailbox)

  print "Opening mbx: [%s]" % mbxfile

  if not os.path.exists(mbxfile):
    return []

  fd = open(mbxfile)

  uidls=[]

  for line in fd.readlines():
    if line[0:7] == "* UIDL:":
      list = line.split(':')
      uidls.append( list[1].strip() )

  fd.close()

  return uidls

  """This function returns a list of all of the uidl (unique id's) of
     all of the messages on the server """

############################################################
def decode_field(field):
  """This function is used for decoding the values of message fields
     (To:, Cc:, Subject:, ..) using python's email.Header.decode_header
     function."""
  field = field.replace('\r\n','')
  field = field.replace('\n','')

  list = email.Header.decode_header (field)

  decoded = " ".join(["%s" % k for (k,v) in list])

  #print "Decoding [%s] to [%s]" % (field, decoded)

  return decoded
  

############################################################
def string2time (strdate):
  """ This function converts a string from a date that contains
      a string in one of the following formats:
      - %a, %d %b %Y %H:%M:%S
      - %d %b %Y %H:%M:%S
      (since some emails come with a date field in the first format,
       and others come with a date field in the second format)

      TODO:
      - maybe add another try+except to the second time.strptime()
      - use the GMT/.. offset instead of ignoring it
      
  """

  ## these 3 if blocks are used to ignore the GMT/.. offset
  if not strdate.find('(') == -1:
    strdate = strdate[0:strdate.find('(')].strip()

  if not strdate.find('+') == -1:
    strdate = strdate[0:strdate.find('+')].strip()

  if not strdate.find('-') == -1:
    strdate = strdate[0:strdate.find('-')].strip()

  ## convert the date string into a  9-item tuple
  try:
    t = time.strptime(strdate, "%a, %d %b %Y %H:%M:%S")
  except ValueError, e:
    t = time.strptime(strdate, "%d %b %Y %H:%M:%S")

  return t
    
############################################################
def buildTitle (tm, mailFrom, mailSubject):
  """ This functions receives 3 parameters:
        tm =  a 9-item tuple that represents the email's date and time
        mailFrom = the contents of the "From:" field
        mailSubject = the contents of the "Subject:" field 

      and it returns a title for the email message (used with vim's folding)
      the title is in this format:
        * Date and Time | From           | Subject

      the width of "From" is 15 characters by default, it's defined by the
      from_column_width variable
  """
  from_column_width = 15

  # Create Time String
  strtime = time.strftime("%d/%m/%Y %H:%M",tm)

  # Create From String
  strfrom = mailFrom

  if len(strfrom) > from_column_width:
    strfrom=strfrom[0:from_column_width]

  strfrom = strfrom.center(from_column_width)

  # Create Subject String
  title = "* %s | %s | %s\r\n" % (strtime, strfrom, mailSubject)

  return title


############################################################
def handleMsg(mailbox, msg, is_subpart=False, strdate=""):
  """ This function handles a message object recursively, it has 
      several tasks:
      - save all of the attachments in the message
      - extract all of the text information into the message body
      - if the email contains html messages they will be converted 
        into text and added to the message body
      - extract all of the field information (To, Cc, From, ...)
        from the message objects
      
  """
  global text
  global attachments
  global fieldFrom, fieldSubject, fieldTime

  # Message/RFC822 parts are bundled this way ==============
  while isinstance(msg.get_payload(),email.Message.Message):
    msg=msg.get_payload()

  if not is_subpart:
    fieldFrom = ""
    fieldSubject = ""
    fieldTime = None    # fieldTime is a 9-item tuple
    text = ""           # the text contents of a message
    attachments = ""

  ## Set the "From" Field ==================================
  if fieldFrom == "" and msg['From'] != None:
    text += "To: %s\n" % decode_field(msg['To'])
    if msg['Cc'] != None:
      text += "Cc: %s\n" % decode_field(msg['Cc'])
    if msg['Bcc'] != None:
      text += "Bcc: %s\n" % decode_field(msg['Bcc'])
    text += "From: %s\n" % decode_field(msg['From'])
    fieldFrom = decode_field(msg['From'])

  ## Set the "Subject" Field ===============================
  if fieldSubject == "" and msg['Subject'] != None:
    fieldSubject = decode_field(msg['Subject'])
    text += "Subject: %s\n" % fieldSubject

  ## Set the "Date" Field ==================================
  if fieldTime == None and msg['Date'] != None:
    fieldTime = string2time(msg['Date'])
    strdate = time.strftime("%Y%m%d%H%M", fieldTime)

  ## Handle multipart messages recursively =================
  if msg.is_multipart():
    for submsg in msg.get_payload():
      handleMsg(mailbox, submsg, True, strdate)
  else:
    fname = msg.get_filename()
    if fname == None:
      if msg.get_content_type() == 'text/plain':
        text += "\n%s" % msg.get_payload(decode=1)
      else:
        fname = "message.htm"

    ## Save an attachment to a file ========================
    if not fname == None:
      fname = decode_field(fname)
      filename = "%s\\att_%s\\%s_%s" % (mailboxdir, mailbox, strdate, fname)
      org_filename = filename
      i = 1
      while os.path.exists(filename):
        path, ext = os.path.splitext(org_filename)
        filename = "%s (%d)%s" % (path, i, ext)
        i = i + 1

      print " Found part: %s" % filename  # for debugging purposes
      attachments += "%s\n" % filename
      fd = open (filename, "wb")
      data = msg.get_payload(decode=1)
      fd.write(data)

      # convert an html message to text
      if fname == "message.htm":
        try:
          strio = cStringIO.StringIO()
          html2text.html2text_file(data, out=strio.write)
          text += strio.getvalue()
          strio.close()
        except sgmllib.SGMLParseError, e:
          print e

      fd.close()

  # if this is the toplevel message (the first function that was called by
  # fetch_mailbox, then return the title of the message
  if not is_subpart and fieldTime != None:
    title = buildTitle(fieldTime, fieldFrom, fieldSubject)
    return title

############################################################
def fetch_mailbox((mailbox, host, user, passwd)):
  """ This function gets an account object as input, this object is a 4-item
      tuple that contains the mailbox name, the address of the pop3 server,
      the username to use and the password.

      What it does is download all of the messages in the server, save the 
      attachments to the $mailbox/attach directory, and save all of the email 
      messages into one .mbx file.

      This function returns the amount of new messages found

      TODO:
      - find a way to encrypt "passwd" instead of asking it from the user 
        each time.

  """

  global text, attachments

  ## login to the pop3 server ==============================
  print
  print "###### Connecting to %s" % host
  M = poplib.POP3(host)
  M.set_debuglevel(1)
  M.user(user)
  M.pass_(getpass.getpass())

  ## create the mailbox and attachments directories if required
  if not os.path.exists (mailboxdir):
    print "Creating Directory %s", mailboxdir
    os.mkdir (mailboxdir)

  att_dir = "%s\\att_%s" % (mailboxdir, mailbox)
  if not os.path.exists (att_dir):
    print "Creating Directory %s", att_dir
    os.mkdir (att_dir)

  
  ## get list of uidls in the mailbox file =================
  uidls = get_mailbox_uidls(mailbox)

  ## get number of messages ================================
  numMessages = len(M.list()[1])
  print "There are %d messages on the server" % numMessages


  ## get uidls from server and compare with the uidls in the
  ## mailbox ===============================================
  uidls_srv = M.uidl()
  list = uidls_srv[1]
  fetchlist = []
  for item in list:
    msgno, uidl = item.split(' ')
    msgno = int(msgno)
    if not uidl in uidls:
      print "Found new message: (%d, %s)" % (msgno, uidl)
      fetchlist.append(msgno)

  print "There are %d new messages on the server" % len(fetchlist)

  alltext = "" ## this variable contains the mbox contents

  ## go over all of the emails =============================
  for i in fetchlist:

    flatmsg = ""

    ## retreive message
#     for line in M.retr(i+1)[1]:
    for line in M.retr(i)[1]:
      flatmsg += line + "\r\n"

    ## parse message
    msg = email.message_from_string (flatmsg)

    ## handle Email.message object
    title = handleMsg(mailbox, msg)


    msgtext = "%s\n%s* UIDL: %s\n%s\n\n" % (''.center(70,'#'), title, uidl, text)
    if not attachments == "":
      msgtext += "#### Attachments:\n%s" % attachments

    alltext = msgtext.replace('\r\n','\n') + alltext

  ## add 'alltext' to the beginning of the mailbox file ====
  mboxfile = "%s\\%s.mbx" % (mailboxdir, mailbox)
  contents = ""
  if os.path.exists(mboxfile):
    mbox = open(mboxfile, "rt")
    contents = mbox.read()
    mbox.close()

  mbox = open(mboxfile, "wt")
  mbox.write (alltext)
  if contents != "":
    mbox.write (contents)

  mbox.close()

  return len(fetchlist)

############################################################

if __name__ == "__main__":
  summary = []
  for account in accounts:
    newMessages = fetch_mailbox (account)
    if newMessages > 0:
      summary.append((account[0], newMessages))

  if len(summary) > 0:
    print
    print "Summary: "
      
    for account,msgs in summary:
      print "  %s: %d new messages" % (account,msgs)
