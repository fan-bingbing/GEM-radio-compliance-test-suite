import smtplib
from email.message import EmailMessage


def SendEmail(Subject, From, To, Timestamp):
    msg = EmailMessage()

    #with open('C:\\Users\\Aaron Fan\\Documents\\GEM-radio-compliance-test-suite\\bottomlevel\\EmailContent.txt') as fp1:
    msg.set_content(f"Test completed at {Timestamp}")


    with open('C:\\Users\\Aaron Fan\\Documents\\GEM-radio-compliance-test-suite\\scripts\\CM60_Result.xlsx', 'rb') as fp2:
        content = fp2.read()
        msg.add_attachment(content, maintype='application', subtype='xlsx', filename='CM60_Result.xlsx')

    msg['subject'] = Subject
    msg['From'] = From
    msg['To'] = To
    #msg['Cc']

    s = smtplib.SMTP('smtp.live.com', 587)# SMTP server used for sending
    s.starttls()# puts connection to SMTP server in TLS mode
    s.login("fanyu1980@hotmail.com", "Password782578")
    s.send_message(msg)
    s.quit()
