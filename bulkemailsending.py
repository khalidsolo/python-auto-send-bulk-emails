#Import necessary packages
import pandas as p
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders



#Load and read exel file from which we get our targeted email ids to which email to be sent
data = p.read_excel('emails.xlsx')
email_column = data.get("email")
name_column = data['name']



#Loading the file which going to be attached with email
message = MIMEMultipart()
filename = "file.pdf"
attachment = open(filename, 'rb')
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('content-Disposition', f'attachment; filename={filename}')
message.attach(part)


#logic used to send emails one-by-one automatically
try:
    for i in range(len(email_column)):
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login("YourGmailId@gmail.com", "YourGmailIdPassword")
        message['subject'] = "Test mail"
        message['To'] = email_column[i]
        # message['bcc'] = email_column[i]
        # message['cc'] = email_column[i]
        body = f"Hi {name_column[i]},\n \t\tHello how r you"
        message.attach(MIMEText(body,'plain')) #'plain means message text is not in html format, for html format write 'html'
        message.as_string()
        server.sendmail('YourGmailId@gmail.com', message['To'],message.as_string())


        #Making Fresh attachments(for next iteration of loop)
        message=MIMEMultipart()
        message.attach(part)
        print("sent"+str(i+1))



    print('all done, pack your Bags')
    server.quit()
except Exception as e:
    print(e)