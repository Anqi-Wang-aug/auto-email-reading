from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr 
import email
import time
import imaplib
import poplib
import re

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def get_att(msg, fpath):
    PATTERN =  r'[^A-Za-z0-9\. ]'
    attachment_files = []
    for part in msg.walk():
        file_name = part.get_filename()
        contType = part.get_content_type()

        if file_name:
            h = email.header.Header(file_name)
            dh = email.header.decode_header(h)
            filename = dh[0][0]
            if dh[0][1]:
                filename = decode_str(str(filename,dh[0][1]))
            print('Downloading')
            data = part.get_payload(decode=True)
            attachment_files.append(filename)
            print('writing')
            filename = re.sub(PATTERN, "", filename)
            print(filename)
            with open(fpath + filename, 'wb') as att_file:  
                att_file.write(data)           
            print('Download Complete')
        else: print('Not attachment')
    return attachment_files


my_email = input('Enter your email: ')
password = input('Password: ')          
pop3_server = 'pop.outlook.com'  

PATH = 'C:\\Users\\Huawei\\Downloads\\'



imap = imaplib.IMAP4_SSL('imap.outlook.com')
print('Connection established')
result = imap.login(my_email, password)
print(result)
imap.select('INBOX')
print('Inbox opened')
response, messages = imap.search(None, 'UNSEEN')
print('Got all Unread files')
msg_list = messages[0].split()
print(len(msg_list))
server = poplib.POP3_SSL(pop3_server)
print(server.getwelcome().decode('utf-8'))
server.user(my_email)
server.pass_(password)

for m in msg_list:
    print(m)
    m=m.decode('utf-8')
    print('Fetched message')
    print('Decoded complete')
    print('Getting Attachments')
    resp, lines, octets = server.retr(m)
    msg_content = b'\r\n'.join(lines).decode('utf-8')
    msg = Parser().parsestr(msg_content)
    get_att(msg, PATH)
    imap.store(m, '+FLAGS', '\\Read')

imap.close()
imap.logout()
