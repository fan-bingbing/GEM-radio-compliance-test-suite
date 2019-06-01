import smtplib
from email.message import EmailMessage

msg = EmailMessage()

with open('C:\\Users\\YU FAN\\Desktop\\pynotes.txt') as fp1:
    msg.set_content(fp1.read())


with open('C:\\Users\\YU FAN\\Desktop\\Book1.xlsx', 'rb') as fp2:
    content = fp2.read()
    msg.add_attachment(content, maintype='application', subtype='xlsx', filename='Book1.xlsx')

msg['subject'] = 'hehe'
msg['From'] = 'fanyu1980@hotmail.com'
msg['To'] = 'fansunsun@163.com'
#msg['Cc']

s = smtplib.SMTP('smtp.live.com', 587)# SMTP server used for sending
s.starttls()# puts connection to SMTP server in TLS mode
s.login("fanyu1980@hotmail.com", "Password782578")
s.send_message(msg)
s.quit()
