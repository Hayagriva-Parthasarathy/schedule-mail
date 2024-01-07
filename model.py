
import os
import ssl
from email.message import EmailMessage
from datetime import datetime
import smtplib
import schedule
import time

class Mail:

   

    

    def mailing(self):
    # Get the current date and time
        current_datetime = datetime.now()

        # Extract the current date
        current_date = current_datetime.date()
        path = r"C:\Users\imvin\Downloads\Book1.xlsx"

        email_sender = 'hayagriva02@gmail.com'
        email_pass = os.environ.get('py_pass')
        email_receiver = ["hayagrivadhoni@gmail.com","hp2461@srmist.edu.in"]
        subject = "Excel sheet for the day"
        body = f"Player stats as on {current_date}"


        for person in email_receiver:

            em = EmailMessage()
            em['From'] = email_sender
            em['To'] = person
            em['Subject'] = subject
            em.set_content(body)
            with open(path, 'rb') as excel_attachment:
                em.add_attachment(
                    excel_attachment.read(),
                    maintype='application',#xl sheet belongs to maintype application
                    subtype='vnd.ms-excel',#all xl files belongs to this subtype
                    filename='stats.xlsx'#renamed filename for mailing
                )

            # Connect to the SMTP server and send the email
            context = ssl.create_default_context()

            # Log in and send the email
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
                smtp.login(email_sender, email_pass)
                smtp.sendmail(email_sender, person, em.as_string())
class Scheduling:
    
    def __init__(self):
        self.mail = Mail()
    

    

sc= Scheduling()
schedule.every().day.at("12:41").do(sc.mail.mailing)


while True:
    schedule.run_pending()
    time.sleep(60)