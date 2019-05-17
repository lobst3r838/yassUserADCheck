from ldap3 import Server, Connection, ALL, NTLM, ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES
from ldap3.core.exceptions import LDAPCursorError
import time

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

server_name = 'IP_OR_FQDN'
domain_name = 'DOMAIN_NAME'
user_name = 'USERNAME'
password = 'PASSWORD'

if __name__ == '__main__':
    dataExport = time.strftime('%Y-%m-%d')

    f = open('userReport_' + str(dataExport) + '.csv', 'w')
    f.write('NOME;USERNAME;NR LOGIN;ULTIMO LOGIN;ULTIMO CAMBIO PASSWORD;USER ACCOUNT CONTROL\n')
    server = Server(server_name, get_info=ALL)
    conn = Connection(server, user='{}\\{}'.format(domain_name, user_name), password=password, authentication=NTLM,
                      auto_bind=True)
    conn.search('dc={},dc=DOMAIN_SUFFIX'.format(domain_name), '(objectcategory=person)',
                attributes=[ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES])
    for e in conn.entries:
        try:
            logonCount = e.logonCount
        except LDAPCursorError:
            logonCount = ""

        try:
            lastLogon = e.lastLogonTimestamp
        except LDAPCursorError:
            lastLogon = ""

        try:
            pwdLastSet = e.pwdLastSet
        except LDAPCursorError:
            pwdLastSet = ""

        try:
            sAMAccountName = e.sAMAccountName
        except LDAPCursorError:
            sAMAccountName = ""

        try:
            userAccountControl = e.userAccountControl

            if userAccountControl == 66048:
                userAccountControl = 'PASSWORD NEVER EXPIRE'
            elif userAccountControl == 66050 or userAccountControl == 514:
                userAccountControl = 'DISABILITATO'
            elif userAccountControl == 512:
                userAccountControl = 'PASSWORD EXPIRE'
            else:
                userAccountControl = 'CONTROLLARE'

        except LDAPCursorError:
            userAccountControl = "CONTATTO?"

        f.write(str(e.name) + ';' + str(sAMAccountName) + ';' + str(logonCount) + ';' + str(lastLogon)[:19] + ';' +
                str(pwdLastSet)[:19] + ';' + str(userAccountControl) + '\n')

    f.close()

    # invio mail
    fromaddr = "SENDER_EMAIL"
    toaddr = "RECIPIENT_EMAIL"

    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = "SUBJECT"
    body = "BODY"

    msg.attach(MIMEText(body, 'plain'))
    filename = 'userReport_' + str(dataExport) + '.csv'
    attachment = open('userReport_' + str(dataExport) + '.csv', 'rb')
    p = MIMEBase('application', 'octet-stream')
    p.set_payload((attachment).read())
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(p)

    s = smtplib.SMTP('SMTP_IP_OR_FQDN', 25)
    text = msg.as_string()
    s.sendmail(fromaddr, toaddr, text)

    s.quit()
