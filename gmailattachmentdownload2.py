# Make sure you have IMAP enabled in your gmail settings.
# Right now it won't download same file name twice even if their contents are different.

import email
import getpass, imaplib
import os
import sys
import datetime
import time
import shutil

detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

#userName = raw_input('Username')
#passwd = getpass.getpass('Passcode')

userName = 'david.balla.fleetmatics'
passwd = 'Fleet2015'


try:
    imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
    typ, accountDetails = imapSession.login(userName, passwd)
    if typ != 'OK':
        print 'Not able to sign in!'
        raise
    
    imapSession.select('[Gmail]/All Mail')
    typ, data = imapSession.search(None, 'ALL')
    if typ != 'OK':
        print 'Error searching Inbox.'
        raise
    
    # Iterating over all emails
    for msgId in data[0].split():
        typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
        if typ != 'OK':
            print 'Error fetching mail.'
            raise

        emailBody = messageParts[0][1]
        mail = email.message_from_string(emailBody)
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                # print part.as_string()
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue
            fileName = part.get_filename()

            if bool(fileName):
                filePath = os.path.join(detach_dir, 'attachments', fileName)
                if not os.path.isfile(filePath) :
                    print fileName
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
    imapSession.close()
    imapSession.logout()
except :
    print 'Not able to download all attachments.'


time.sleep(5) # delays for 5 seconds

m = imaplib.IMAP4_SSL("imap.gmail.com")  # server to connect to
print "Connecting to mailbox..."
print time.ctime()
m.login('david.balla.fleetmatics@gmail.com', 'Fleet2015')

print m.select('[Gmail]/All Mail')  # required to perform search, m.list() for all lables, '[Gmail]/Sent Mail'

before_date = (datetime.date.today() - datetime.timedelta(-1)).strftime("%d-%b-%Y")  # date string, 04-Jan-2013
print before_date
typ, data = m.search(None, '(BEFORE {0})'.format(before_date))  # search pointer for msgs before before_date

if data != ['']:  # if not empty list means messages exist
    no_msgs = data[0].split()[-1]  # last msg id in the list
    print "To be removed:\t", no_msgs, "messages found with date before", before_date
    m.store("1:{0}".format(no_msgs), '+X-GM-LABELS', '\\Trash')  # move to trash
    print "Deleted {0} messages. Closing connection & logging out.".format(no_msgs)
else:
    print "Nothing to remove."

#This block empties trash, remove if you want to keep, Gmail auto purges trash after 30 days.
print("Emptying Trash & Expunge...")
m.select('[Gmail]/Trash')  # select all trash
m.store("1:*", '+FLAGS', '\\Deleted')  #Flag all Trash as Deleted
m.expunge()  # not need if auto-expunge enabled

print("Done. Closing connection & logging out.")
m.close()
m.logout()
print "All Done."


path = '/home/pi/attachments/'
print path
filelist1 = os.listdir(path)[0]
print filelist1
orgfile = path + filelist1
path1 = '/home/pi/working/'
print path1
shutil.move(orgfile, path1)
#os.remove(orgfile)
execfile('FinishQuestionsFinal.py')
#execfile('DeleteFinish.py')


